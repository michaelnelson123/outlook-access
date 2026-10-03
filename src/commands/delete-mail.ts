// src/commands/delete-mail.ts
//
// Soft-delete one or more messages: a move to the well-known DeletedItems
// folder, issued through move-mail's own code path (one up-front resolve of
// the destination, then one `POST /me/messages/{id}/move` per id). There is
// deliberately no hard delete — `DELETE /me/messages/{id}` and a move to
// RecoverableItemsDeletions are both out of scope, so anything this command
// removes can be recovered from Deleted Items.
//
// Output mirrors move-mail with `moved` renamed to `deleted`: each entry
// carries the source id and the NEW id the message has in Deleted Items
// (a move re-keys the message). `--continue-on-error` and the exit-5 rule
// for a non-empty `failed[]` are move-mail's.
//
// `--dry-run` validates argv and reports what would be moved without
// loading the session or contacting M365.

import type { MoveDestination, MoveEntry, MoveFailedEntry } from '../folders/types';

import { UsageError } from './list-mail';
import { run as runMoveMail, type MoveMailDeps } from './move-mail';

export type DeleteMailDeps = MoveMailDeps;

/** Well-known alias every soft delete targets. */
export const DELETE_DESTINATION = 'DeletedItems';

export interface DeleteMailOptions {
  /** If true, per-message failures are collected into `failed[]` instead of short-circuiting. */
  continueOnError?: boolean;
  /** Validate and report the planned move; do not contact M365. */
  dryRun?: boolean;
}

export interface DeleteMailResult {
  /** `deleted` when moves were issued; `dry-run` when nothing was sent. */
  mode: 'deleted' | 'dry-run';
  /** Resolved Deleted Items folder; `null` under dry-run (no lookup is made). */
  destination: MoveDestination | null;
  /**
   * Messages moved to Deleted Items with their new ids. Under dry-run, the
   * ids that would be moved, with no `newId`.
   */
  deleted: (MoveEntry | { sourceId: string })[];
  failed: MoveFailedEntry[];
  summary: { requested: number; deleted: number; failed: number };
}

export async function run(
  deps: DeleteMailDeps,
  messageIds: string[],
  opts: DeleteMailOptions = {},
): Promise<DeleteMailResult> {
  if (!Array.isArray(messageIds) || messageIds.length === 0) {
    throw new UsageError('delete-mail: at least one <messageId> positional argument is required');
  }
  for (const id of messageIds) {
    if (typeof id !== 'string' || id.length === 0) {
      throw new UsageError(
        'delete-mail: <messageId> positional arguments must be non-empty strings',
      );
    }
  }

  if (opts.dryRun === true) {
    return {
      mode: 'dry-run',
      destination: null,
      deleted: messageIds.map((sourceId) => ({ sourceId })),
      failed: [],
      summary: { requested: messageIds.length, deleted: messageIds.length, failed: 0 },
    };
  }

  const moved = await runMoveMail(deps, messageIds, {
    to: DELETE_DESTINATION,
    continueOnError: opts.continueOnError === true,
  });

  return {
    mode: 'deleted',
    destination: moved.destination,
    deleted: moved.moved,
    failed: moved.failed,
    summary: {
      requested: moved.summary.requested,
      deleted: moved.summary.moved,
      failed: moved.summary.failed,
    },
  };
}
