// test_scripts/commands-mark-delete-mail.spec.ts
//
// Command-level tests for `mark-mail` and `delete-mail`, plus get-mail's
// plain-text body request.
//
// Scope:
//   - mark-mail: exactly one action flag (else UsageError → exit 2); the
//     PATCH body per action; sequential loop; --continue-on-error collects
//     failures; first failure aborts otherwise; --dry-run sends nothing.
//   - delete-mail: a move to the DeletedItems alias through move-mail's
//     path (one resolve, then one move per id); new ids surfaced in
//     deleted[]; --continue-on-error; --dry-run sends nothing.
//   - get-mail: --body text passes bodyContentType 'text' to client.get;
//     html / none do not.
//
// No real HTTP. The OutlookClient is mocked with `Partial<OutlookClient>`.

import { describe, expect, it, vi } from 'vitest';

import { run as runDeleteMail } from '../src/commands/delete-mail';
import { run as runGetMail } from '../src/commands/get-mail';
import { UsageError } from '../src/commands/list-mail';
import { run as runMarkMail } from '../src/commands/mark-mail';
import type { MarkMailDeps } from '../src/commands/mark-mail';
import type { CliConfig } from '../src/config/config';
import { ApiError } from '../src/http/errors';
import type { OutlookClient } from '../src/http/outlook-client';
import type { FolderSummary, MessageSummary } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const SESSION: SessionFile = {
  version: 1,
  capturedAt: '2026-04-21T12:00:00.000Z',
  account: { upn: 'a@b', puid: 'p', tenantId: 't' },
  bearer: {
    token: 'x.y.z',
    expiresAt: '2099-04-21T12:00:00.000Z',
    audience: 'https://outlook.office.com',
    scopes: [],
  },
  cookies: [],
  anchorMailbox: 'PUID:p@t',
};

const CONFIG = {
  httpTimeoutMs: 5_000,
  loginTimeoutMs: 60_000,
  chromeChannel: 'chrome',
  sessionFilePath: '/tmp/session.json',
  profileDir: '/tmp/profile',
  tz: 'UTC',
  outputMode: 'json',
  listMailTop: 10,
  listMailFolder: 'Inbox',
  bodyMode: 'text',
  calFrom: 'now',
  calTo: 'now + 7d',
  quiet: true,
  noAutoReauth: true,
} as unknown as CliConfig;

function buildDeps(client: Partial<OutlookClient>): MarkMailDeps & {
  loadSession: ReturnType<typeof vi.fn>;
} {
  return {
    config: CONFIG,
    sessionPath: '/tmp/session.json',
    loadSession: vi.fn(async () => SESSION),
    saveSession: async () => {
      /* no-op */
    },
    doAuthCapture: async () => SESSION,
    createClient: () => client as OutlookClient,
  };
}

function notFound(id: string): ApiError {
  return new ApiError({
    code: 'NOT_FOUND',
    message: `message ${id} is gone`,
    httpStatus: 404,
    url: `https://outlook.office.com/api/v2.0/me/messages/${id}`,
  });
}

function deletedItemsFolder(): FolderSummary {
  return { Id: 'deleted-raw-id', DisplayName: 'Deleted Items', WellKnownName: 'deleteditems' };
}

function moved(newId: string): MessageSummary {
  return {
    Id: newId,
    Subject: 'moved',
    ReceivedDateTime: '2026-04-21T09:00:00Z',
    HasAttachments: false,
    IsRead: true,
    WebLink: '',
  };
}

// ---------------------------------------------------------------------------
// mark-mail
// ---------------------------------------------------------------------------

describe('mark-mail command', () => {
  it.each([
    ['read', { IsRead: true }],
    ['unread', { IsRead: false }],
    ['flag', { Flag: { FlagStatus: 'Flagged' } }],
    ['unflag', { Flag: { FlagStatus: 'NotFlagged' } }],
    ['complete', { Flag: { FlagStatus: 'Complete' } }],
  ])('--%s PATCHes every id with %j', async (action, patch) => {
    const updateMessage = vi.fn(async () => ({ Id: 'x', Subject: '' }));
    const deps = buildDeps({ updateMessage });

    const result = await runMarkMail(deps, ['m1', 'm2'], { [action]: true });

    expect(updateMessage).toHaveBeenCalledTimes(2);
    expect(updateMessage).toHaveBeenNthCalledWith(1, 'm1', patch);
    expect(updateMessage).toHaveBeenNthCalledWith(2, 'm2', patch);
    expect(result).toEqual({
      mode: 'applied',
      action,
      patch,
      marked: [{ id: 'm1' }, { id: 'm2' }],
      failed: [],
      summary: { requested: 2, marked: 2, failed: 0 },
    });
  });

  it('no action flag raises UsageError', async () => {
    const deps = buildDeps({ updateMessage: vi.fn() });
    await expect(runMarkMail(deps, ['m1'], {})).rejects.toBeInstanceOf(UsageError);
  });

  it('two action flags raise UsageError and nothing is loaded or sent', async () => {
    const updateMessage = vi.fn();
    const deps = buildDeps({ updateMessage });
    await expect(runMarkMail(deps, ['m1'], { read: true, flag: true })).rejects.toBeInstanceOf(
      UsageError,
    );
    expect(deps.loadSession).not.toHaveBeenCalled();
    expect(updateMessage).not.toHaveBeenCalled();
  });

  it('empty id list raises UsageError', async () => {
    const deps = buildDeps({ updateMessage: vi.fn() });
    await expect(runMarkMail(deps, [], { read: true })).rejects.toBeInstanceOf(UsageError);
  });

  it('first failure aborts without --continue-on-error', async () => {
    const updateMessage = vi.fn(async (id: string) => {
      if (id === 'm1') throw notFound(id);
      return { Id: id, Subject: '' };
    });
    const deps = buildDeps({ updateMessage });
    await expect(runMarkMail(deps, ['m1', 'm2'], { unread: true })).rejects.toBeDefined();
    expect(updateMessage).toHaveBeenCalledTimes(1);
  });

  it('--continue-on-error collects failures and carries on', async () => {
    const updateMessage = vi.fn(async (id: string) => {
      if (id === 'm2') throw notFound(id);
      return { Id: id, Subject: '' };
    });
    const deps = buildDeps({ updateMessage });

    const result = await runMarkMail(deps, ['m1', 'm2', 'm3'], {
      complete: true,
      continueOnError: true,
    });

    expect(updateMessage).toHaveBeenCalledTimes(3);
    expect(result.marked).toEqual([{ id: 'm1' }, { id: 'm3' }]);
    expect(result.failed).toHaveLength(1);
    expect(result.failed[0].id).toBe('m2');
    expect(result.failed[0].error.httpStatus).toBe(404);
    expect(result.failed[0].error.code.length).toBeGreaterThan(0);
    expect(result.summary).toEqual({ requested: 3, marked: 2, failed: 1 });
  });

  it('--dry-run reports the plan without loading the session or sending', async () => {
    const updateMessage = vi.fn();
    const deps = buildDeps({ updateMessage });

    const result = await runMarkMail(deps, ['m1'], { flag: true, dryRun: true });

    expect(deps.loadSession).not.toHaveBeenCalled();
    expect(updateMessage).not.toHaveBeenCalled();
    expect(result.mode).toBe('dry-run');
    expect(result.patch).toEqual({ Flag: { FlagStatus: 'Flagged' } });
    expect(result.marked).toEqual([{ id: 'm1' }]);
  });
});

// ---------------------------------------------------------------------------
// delete-mail
// ---------------------------------------------------------------------------

describe('delete-mail command', () => {
  it('resolves DeletedItems once, moves every id there and reports new ids', async () => {
    const getFolder = vi.fn(async (alias: string) => {
      if (alias === 'DeletedItems') return deletedItemsFolder();
      throw new Error(`unexpected getFolder(${alias})`);
    });
    const moveMessage = vi.fn(async (id: string) => moved(`${id}-NEW`));
    const deps = buildDeps({ getFolder, moveMessage });

    const result = await runDeleteMail(deps, ['m1', 'm2']);

    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith('DeletedItems');
    expect(moveMessage).toHaveBeenNthCalledWith(1, 'm1', 'deleted-raw-id');
    expect(moveMessage).toHaveBeenNthCalledWith(2, 'm2', 'deleted-raw-id');
    expect(result.mode).toBe('deleted');
    expect(result.destination?.Id).toBe('deleted-raw-id');
    expect(result.deleted).toEqual([
      { sourceId: 'm1', newId: 'm1-NEW' },
      { sourceId: 'm2', newId: 'm2-NEW' },
    ]);
    expect(result.failed).toEqual([]);
    expect(result.summary).toEqual({ requested: 2, deleted: 2, failed: 0 });
  });

  it('--continue-on-error collects failures in failed[]', async () => {
    const getFolder = vi.fn(async () => deletedItemsFolder());
    const moveMessage = vi.fn(async (id: string) => {
      if (id === 'm1') throw notFound(id);
      return moved(`${id}-NEW`);
    });
    const deps = buildDeps({ getFolder, moveMessage });

    const result = await runDeleteMail(deps, ['m1', 'm2'], { continueOnError: true });

    expect(result.deleted).toEqual([{ sourceId: 'm2', newId: 'm2-NEW' }]);
    expect(result.failed[0].sourceId).toBe('m1');
    expect(result.failed[0].error.httpStatus).toBe(404);
    expect(result.summary).toEqual({ requested: 2, deleted: 1, failed: 1 });
  });

  it('first failure aborts without --continue-on-error', async () => {
    const getFolder = vi.fn(async () => deletedItemsFolder());
    const moveMessage = vi.fn(async (id: string) => {
      throw notFound(id);
    });
    const deps = buildDeps({ getFolder, moveMessage });
    await expect(runDeleteMail(deps, ['m1', 'm2'])).rejects.toBeDefined();
    expect(moveMessage).toHaveBeenCalledTimes(1);
  });

  it('empty id list raises UsageError', async () => {
    const deps = buildDeps({});
    await expect(runDeleteMail(deps, [])).rejects.toBeInstanceOf(UsageError);
  });

  it('--dry-run reports the plan without loading the session or sending', async () => {
    const getFolder = vi.fn();
    const moveMessage = vi.fn();
    const deps = buildDeps({ getFolder, moveMessage });

    const result = await runDeleteMail(deps, ['m1'], { dryRun: true });

    expect(deps.loadSession).not.toHaveBeenCalled();
    expect(getFolder).not.toHaveBeenCalled();
    expect(moveMessage).not.toHaveBeenCalled();
    expect(result).toEqual({
      mode: 'dry-run',
      destination: null,
      deleted: [{ sourceId: 'm1' }],
      failed: [],
      summary: { requested: 1, deleted: 1, failed: 0 },
    });
  });
});

// ---------------------------------------------------------------------------
// get-mail --body text
// ---------------------------------------------------------------------------

describe('get-mail --body', () => {
  function clientWithGet() {
    const get = vi.fn(async (path: string) =>
      path.endsWith('/attachments') ? { value: [] } : { Id: 'm1', Body: { ContentType: 'Text' } },
    );
    return { get };
  }

  it('text asks for a plain-text body on the message GET only', async () => {
    const client = clientWithGet();
    await runGetMail(buildDeps(client), 'm1', { body: 'text' });
    const messageCall = client.get.mock.calls.find((c) => !String(c[0]).endsWith('/attachments'));
    const attachCall = client.get.mock.calls.find((c) => String(c[0]).endsWith('/attachments'));
    expect(messageCall?.[2]).toEqual({ bodyContentType: 'text' });
    expect(attachCall?.[2]).toBeUndefined();
  });

  it.each(['html', 'none'] as const)('%s does not ask for plain text', async (body) => {
    const client = clientWithGet();
    await runGetMail(buildDeps(client), 'm1', { body });
    for (const call of client.get.mock.calls) {
      expect(call[2]).toBeUndefined();
    }
  });
});
