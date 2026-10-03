// src/commands/get-mail.ts
//
// Retrieve a single message + its attachment metadata.
// See project-design.md §2.13.4 and refined spec §5.4.

import type { CliConfig } from '../config/config';
import {
  buildExtendedPropertiesExpand,
  ExtendedPropertyError,
  parseExtendedPropertyIds,
} from '../http/extended-properties';
import type { OutlookClient } from '../http/outlook-client';
import type { AttachmentSummary, Message, ODataListResponse } from '../http/types';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export interface GetMailDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export type BodyMode = 'html' | 'text' | 'none';

export interface GetMailOptions {
  body?: BodyMode;
  /**
   * MAPI property ids to return under `SingleValueExtendedProperties`, e.g.
   * `Binary 0x348A`. Repeatable; each value may be comma-separated.
   */
  extendedProperty?: string[];
}

const BODY_MODES: readonly BodyMode[] = ['html', 'text', 'none'];

export async function run(
  deps: GetMailDeps,
  id: string,
  opts: GetMailOptions = {},
): Promise<Message> {
  if (typeof id !== 'string' || id.length === 0) {
    throw new UsageError('get-mail: <id> is required');
  }

  const body: BodyMode = opts.body ?? deps.config.bodyMode;
  if (!BODY_MODES.includes(body)) {
    throw new UsageError(
      `get-mail: --body must be one of ${BODY_MODES.join('|')} (got ${String(body)})`,
    );
  }

  let expand: string;
  try {
    expand = buildExtendedPropertiesExpand(parseExtendedPropertyIds(opts.extendedProperty));
  } catch (err) {
    if (err instanceof ExtendedPropertyError) {
      throw new UsageError(`get-mail: ${err.message}`);
    }
    throw err;
  }

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  const encodedId = encodeURIComponent(id);

  try {
    const [message, attachments] = await Promise.all([
      // `--body text` asks the server for a plain-text body via the Prefer
      // header; without it the stored HTML body comes back.
      client.get<Message>(
        `/api/v2.0/me/messages/${encodedId}`,
        expand.length > 0 ? { $expand: expand } : undefined,
        body === 'text' ? { bodyContentType: 'text' } : undefined,
      ),
      client.get<ODataListResponse<AttachmentSummary>>(
        `/api/v2.0/me/messages/${encodedId}/attachments`,
        { $select: 'Id,Name,ContentType,Size,IsInline' },
      ),
    ]);

    const merged: Message = {
      ...message,
      Attachments: Array.isArray(attachments.value) ? attachments.value : [],
    };

    // Body handling. The client does not convert HTML→text itself (ADR
    // deferral in project-design §2.13.4); "text" is converted server-side
    // via the Prefer header above, "none" is stripped here, and "html" is
    // passed through.
    if (body === 'none') {
      delete merged.Body;
    }

    return merged;
  } catch (err) {
    throw mapHttpError(err);
  }
}
