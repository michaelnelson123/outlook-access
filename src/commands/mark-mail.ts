// src/commands/mark-mail.ts
//
// Set the read state or the follow-up flag on one or more messages, via
// `PATCH /api/v2.0/me/messages/{id}` (`OutlookClient.updateMessage`).
//
// Exactly one action per run:
//
//   --read      → { "IsRead": true }
//   --unread    → { "IsRead": false }
//   --flag      → { "Flag": { "FlagStatus": "Flagged" } }
//   --unflag    → { "Flag": { "FlagStatus": "NotFlagged" } }
//   --complete  → { "Flag": { "FlagStatus": "Complete" } }
//
// Conventions mirror move-mail: the loop is strictly sequential, the first
// failure aborts the run unless `--continue-on-error` is set, in which case
// failures are collected into `failed[]` and cli.ts maps a non-empty
// `failed[]` to exit 5 after the result is emitted. Unlike a move, a PATCH
// does not change the message id, so each entry carries just `id`.
//
// `--dry-run` validates argv and reports what would be sent without loading
// the session or contacting M365.

import type { CliConfig } from '../config/config';
import { UpstreamError } from '../config/errors';
import type { OutlookClient, UpdateMessagePatch } from '../http/outlook-client';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

// ---------------------------------------------------------------------------
// Public shapes
// ---------------------------------------------------------------------------

export interface MarkMailDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export type MarkAction = 'read' | 'unread' | 'flag' | 'unflag' | 'complete';

export interface MarkMailOptions {
  read?: boolean;
  unread?: boolean;
  flag?: boolean;
  unflag?: boolean;
  complete?: boolean;
  /** If true, per-message failures are collected into `failed[]` instead of short-circuiting. */
  continueOnError?: boolean;
  /** Validate and report the planned PATCH; do not contact M365. */
  dryRun?: boolean;
}

/** One entry in `MarkMailResult.failed[]`. Same error block as move-mail. */
export interface MarkFailedEntry {
  id: string;
  error: { code: string; httpStatus?: number; message?: string };
}

export interface MarkMailResult {
  /** `applied` when PATCHes were sent; `dry-run` when nothing was sent. */
  mode: 'applied' | 'dry-run';
  action: MarkAction;
  /** The exact PATCH body sent (or that would be sent) for every id. */
  patch: UpdateMessagePatch;
  /** Ids patched successfully (under dry-run: the ids that would be patched). */
  marked: { id: string }[];
  failed: MarkFailedEntry[];
  summary: { requested: number; marked: number; failed: number };
}

const ACTIONS: readonly MarkAction[] = ['read', 'unread', 'flag', 'unflag', 'complete'];

/** PATCH body for each action. */
export function patchFor(action: MarkAction): UpdateMessagePatch {
  switch (action) {
    case 'read':
      return { IsRead: true };
    case 'unread':
      return { IsRead: false };
    case 'flag':
      return { Flag: { FlagStatus: 'Flagged' } };
    case 'unflag':
      return { Flag: { FlagStatus: 'NotFlagged' } };
    case 'complete':
      return { Flag: { FlagStatus: 'Complete' } };
  }
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

export async function run(
  deps: MarkMailDeps,
  messageIds: string[],
  opts: MarkMailOptions = {},
): Promise<MarkMailResult> {
  // ---- argv validation (raises exit 2) ----
  if (!Array.isArray(messageIds) || messageIds.length === 0) {
    throw new UsageError('mark-mail: at least one <messageId> positional argument is required');
  }
  for (const id of messageIds) {
    if (typeof id !== 'string' || id.length === 0) {
      throw new UsageError('mark-mail: <messageId> positional arguments must be non-empty strings');
    }
  }
  const chosen = ACTIONS.filter((a) => opts[a] === true);
  if (chosen.length !== 1) {
    throw new UsageError(
      `mark-mail: exactly one of --read | --unread | --flag | --unflag | --complete is required ` +
        `(got ${chosen.length === 0 ? 'none' : chosen.map((a) => `--${a}`).join(' ')})`,
    );
  }
  const action = chosen[0];
  const patch = patchFor(action);
  const continueOnError = opts.continueOnError === true;

  if (opts.dryRun === true) {
    return {
      mode: 'dry-run',
      action,
      patch,
      marked: messageIds.map((id) => ({ id })),
      failed: [],
      summary: { requested: messageIds.length, marked: messageIds.length, failed: 0 },
    };
  }

  // ---- session + client ----
  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  // ---- per-message loop (sequential, like move-mail) ----
  const marked: { id: string }[] = [];
  const failed: MarkFailedEntry[] = [];

  for (const id of messageIds) {
    try {
      await client.updateMessage(id, patch);
      marked.push({ id });
    } catch (err) {
      const mapped = mapHttpError(err);
      if (continueOnError) {
        failed.push(toFailedEntry(id, mapped));
        continue;
      }
      throw mapped;
    }
  }

  return {
    mode: 'applied',
    action,
    patch,
    marked,
    failed,
    summary: { requested: messageIds.length, marked: marked.length, failed: failed.length },
  };
}

/**
 * Translate an already-mapped error into a `MarkFailedEntry`. Same rules as
 * move-mail's helper: no `cause`, no raw body; `UpstreamError` messages are
 * already redacted.
 */
function toFailedEntry(id: string, err: unknown): MarkFailedEntry {
  if (err instanceof UpstreamError) {
    return { id, error: { code: err.code, httpStatus: err.httpStatus, message: err.message } };
  }
  const maybe = err as { code?: unknown; message?: unknown };
  const code =
    typeof maybe.code === 'string' && maybe.code.length > 0 ? maybe.code : 'UPSTREAM_UNKNOWN';
  const message = typeof maybe.message === 'string' ? maybe.message : String(err);
  return { id, error: { code, message } };
}
