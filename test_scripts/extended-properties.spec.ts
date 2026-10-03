// test_scripts/extended-properties.spec.ts
//
// Tests for `--extended-property` (plan-003): the id parser and $expand
// builder in src/http/extended-properties.ts, the client's pass-through of
// `$expand` on the message-list and folder-list endpoints, and the three
// commands that expose the flag (get-mail, list-mail, list-folders). No real
// HTTP: the client tests stub global fetch, the command tests inject a
// Partial<OutlookClient>.

import { afterEach, describe, expect, it, vi } from 'vitest';

import type { CliConfig } from '../src/config/config';
import { run as runGetMail, type GetMailDeps } from '../src/commands/get-mail';
import { run as runListFolders, type ListFoldersDeps } from '../src/commands/list-folders';
import { run as runListMail, UsageError, type ListMailDeps } from '../src/commands/list-mail';
import {
  buildExtendedPropertiesExpand,
  ExtendedPropertyError,
  parseExtendedPropertyIds,
} from '../src/http/extended-properties';
import { createOutlookClient, type OutlookClient } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';

const DELETED_FROM = 'Binary 0x348A';
const ENTRY_ID = 'Binary 0x0FFF';
const DELETED_FROM_EXPAND = `SingleValueExtendedProperties($filter=PropertyId eq 'Binary 0x348A')`;

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

function buildFakeSession(): SessionFile {
  return {
    version: 1,
    capturedAt: '2026-04-21T12:00:00.000Z',
    account: { upn: 'alice@contoso.com', puid: '1234567890', tenantId: 'tenant-id-abc' },
    bearer: {
      token: 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc',
      expiresAt: '2099-04-21T12:00:00.000Z',
      audience: 'https://outlook.office.com',
      scopes: ['Mail.Read'],
    },
    cookies: [],
    anchorMailbox: 'PUID:1234567890@tenant-id-abc',
  };
}

function buildFakeConfig(): CliConfig {
  return {
    httpTimeoutMs: 5000,
    loginTimeoutMs: 60000,
    chromeChannel: 'chrome',
    sessionFilePath: '/tmp/never-touched.json',
    profileDir: '/tmp/never-touched-profile',
    tz: 'UTC',
    outputMode: 'json',
    listMailTop: 10,
    listMailFolder: 'Inbox',
    bodyMode: 'text',
    calFrom: 'now',
    calTo: 'now + 7d',
    quiet: true,
    noAutoReauth: false,
  };
}

function buildDeps(client: Partial<OutlookClient>): GetMailDeps & ListMailDeps & ListFoldersDeps {
  return {
    config: buildFakeConfig(),
    sessionPath: '/tmp/never-touched.json',
    loadSession: vi.fn(async () => buildFakeSession()),
    saveSession: vi.fn(async () => undefined),
    doAuthCapture: vi.fn(async () => buildFakeSession()),
    createClient: vi.fn(() => client as OutlookClient),
  };
}

function jsonResponse(body: unknown): Response {
  return {
    status: 200,
    ok: true,
    headers: new Headers({ 'content-type': 'application/json' }),
    text: async () => JSON.stringify(body),
    json: async () => body,
  } as unknown as Response;
}

// ---------------------------------------------------------------------------
// parseExtendedPropertyIds / buildExtendedPropertiesExpand
// ---------------------------------------------------------------------------

describe('parseExtendedPropertyIds', () => {
  it('returns [] when the flag is absent', () => {
    expect(parseExtendedPropertyIds(undefined)).toEqual([]);
  });

  it('accepts a tagged property id', () => {
    expect(parseExtendedPropertyIds([DELETED_FROM])).toEqual([DELETED_FROM]);
  });

  it('accepts named properties by Id and by Name', () => {
    const byId = 'String {00020329-0000-0000-C000-000000000046} Id 0x8005';
    const byName = 'String {00020329-0000-0000-C000-000000000046} Name Keywords';
    expect(parseExtendedPropertyIds([byId, byName])).toEqual([byId, byName]);
  });

  it('splits comma-separated values, trims, collapses spaces, and de-duplicates', () => {
    expect(parseExtendedPropertyIds([` ${DELETED_FROM} ,Binary  0x0FFF`, DELETED_FROM])).toEqual([
      DELETED_FROM,
      ENTRY_ID,
    ]);
  });

  it('drops blank entries', () => {
    expect(parseExtendedPropertyIds(['', ' , '])).toEqual([]);
  });

  it.each([
    ['0x348A'],
    ['Bytes 0x348A'],
    ['Binary 348A'],
    ['Binary 0x348'],
    ["Binary 0x348A') or (PropertyId eq 'x"],
    ['String {not-a-guid} Name Keywords'],
    ["String {00020329-0000-0000-C000-000000000046} Name O'Brien"],
  ])('rejects %j', (bad) => {
    expect(() => parseExtendedPropertyIds([bad])).toThrow(ExtendedPropertyError);
  });
});

describe('buildExtendedPropertiesExpand', () => {
  it('returns an empty string for no ids', () => {
    expect(buildExtendedPropertiesExpand([])).toBe('');
  });

  it('filters to one property', () => {
    expect(buildExtendedPropertiesExpand([DELETED_FROM])).toBe(DELETED_FROM_EXPAND);
  });

  it('joins several properties with or', () => {
    expect(buildExtendedPropertiesExpand([DELETED_FROM, ENTRY_ID])).toBe(
      `SingleValueExtendedProperties($filter=PropertyId eq 'Binary 0x348A' or PropertyId eq 'Binary 0x0FFF')`,
    );
  });
});

// ---------------------------------------------------------------------------
// Client: $expand reaches the wire
// ---------------------------------------------------------------------------

describe('outlook-client $expand pass-through', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function clientWithFetch(fetchMock: ReturnType<typeof vi.fn>): OutlookClient {
    vi.stubGlobal('fetch', fetchMock);
    return createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: true,
      onReauthNeeded: async () => buildFakeSession(),
    });
  }

  it('listMessagesInFolder sends $expand when opts.expand is set', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ value: [] }));
    const client = clientWithFetch(fetchMock);
    await client.listMessagesInFolder('DeletedItems', { top: 5, expand: DELETED_FROM_EXPAND });
    const url = new URL(String(fetchMock.mock.calls[0][0]));
    expect(url.searchParams.get('$expand')).toBe(DELETED_FROM_EXPAND);
  });

  it('listMessagesInFolder omits $expand when opts.expand is absent', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ value: [] }));
    const client = clientWithFetch(fetchMock);
    await client.listMessagesInFolder('Inbox', { top: 5 });
    const url = new URL(String(fetchMock.mock.calls[0][0]));
    expect(url.searchParams.has('$expand')).toBe(false);
  });

  it('listFolders sends $expand when given', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ value: [] }));
    const client = clientWithFetch(fetchMock);
    const expand = buildExtendedPropertiesExpand([ENTRY_ID]);
    await client.listFolders('MsgFolderRoot', 100, expand);
    const url = new URL(String(fetchMock.mock.calls[0][0]));
    expect(url.searchParams.get('$expand')).toBe(expand);
  });
});

// ---------------------------------------------------------------------------
// Commands
// ---------------------------------------------------------------------------

describe('get-mail --extended-property', () => {
  it('passes $expand on the message GET and returns the properties', async () => {
    const prop = { PropertyId: DELETED_FROM, Value: 'AAEC' };
    const get = vi.fn(async (path: string) =>
      path.endsWith('/attachments')
        ? { value: [] }
        : { Id: 'm1', Subject: 's', SingleValueExtendedProperties: [prop] },
    );
    const deps = buildDeps({ get } as Partial<OutlookClient>);
    const msg = await runGetMail(deps, 'm1', { extendedProperty: [DELETED_FROM] });
    // Third argument is the RequestOptions slot; this fixture's config sets
    // bodyMode 'text', so the plain-text Prefer header rides along.
    expect(get).toHaveBeenCalledWith(
      '/api/v2.0/me/messages/m1',
      { $expand: DELETED_FROM_EXPAND },
      { bodyContentType: 'text' },
    );
    expect(msg.SingleValueExtendedProperties).toEqual([prop]);
  });

  it('sends no query when the flag is absent (unchanged behaviour)', async () => {
    const get = vi.fn(async (path: string) =>
      path.endsWith('/attachments') ? { value: [] } : { Id: 'm1' },
    );
    const deps = buildDeps({ get } as Partial<OutlookClient>);
    await runGetMail(deps, 'm1');
    expect(get).toHaveBeenCalledWith('/api/v2.0/me/messages/m1', undefined, {
      bodyContentType: 'text',
    });
  });

  it('rejects a malformed id with UsageError before any HTTP', async () => {
    const get = vi.fn();
    const deps = buildDeps({ get } as Partial<OutlookClient>);
    await expect(runGetMail(deps, 'm1', { extendedProperty: ['0x348A'] })).rejects.toBeInstanceOf(
      UsageError,
    );
    expect(get).not.toHaveBeenCalled();
  });
});

describe('list-mail --extended-property', () => {
  it('forwards the expand clause to listMessagesInFolder', async () => {
    const listMessagesInFolder = vi.fn(async () => []);
    const deps = buildDeps({ listMessagesInFolder } as Partial<OutlookClient>);
    await runListMail(deps, { folder: 'DeletedItems', extendedProperty: [DELETED_FROM] });
    const [folderId, opts] = listMessagesInFolder.mock.calls[0] as unknown as [
      string,
      { expand?: string },
    ];
    expect(folderId).toBe('DeletedItems');
    expect(opts.expand).toBe(DELETED_FROM_EXPAND);
  });

  it('forwards it on the --all path too', async () => {
    const listMessagesInFolderAll = vi.fn(async () => ({ messages: [], truncated: false }));
    const deps = buildDeps({ listMessagesInFolderAll } as Partial<OutlookClient>);
    await runListMail(deps, {
      folder: 'DeletedItems',
      all: true,
      extendedProperty: [DELETED_FROM],
    });
    const opts = (
      listMessagesInFolderAll.mock.calls[0] as unknown as [string, { expand?: string }]
    )[1];
    expect(opts.expand).toBe(DELETED_FROM_EXPAND);
  });

  it('leaves expand undefined when the flag is absent', async () => {
    const listMessagesInFolder = vi.fn(async () => []);
    const deps = buildDeps({ listMessagesInFolder } as Partial<OutlookClient>);
    await runListMail(deps, { folder: 'Inbox' });
    const opts = (
      listMessagesInFolder.mock.calls[0] as unknown as [string, { expand?: string }]
    )[1];
    expect(opts.expand).toBeUndefined();
  });

  it('rejects a malformed id with UsageError', async () => {
    const deps = buildDeps({});
    await expect(
      runListMail(deps, { extendedProperty: ['Binary 0x348A,nope'] }),
    ).rejects.toBeInstanceOf(UsageError);
  });
});

describe('list-folders --extended-property', () => {
  it('forwards the expand clause to every listFolders call in a recursive walk', async () => {
    const listFolders = vi.fn(async (parentId: string) =>
      parentId === 'MsgFolderRoot' ? [{ Id: 'f1', DisplayName: 'Inbox', ChildFolderCount: 1 }] : [],
    );
    const deps = buildDeps({ listFolders } as Partial<OutlookClient>);
    await runListFolders(deps, { recursive: true, extendedProperty: [ENTRY_ID] });
    const expand = buildExtendedPropertiesExpand([ENTRY_ID]);
    expect(listFolders).toHaveBeenCalledTimes(2);
    for (const call of listFolders.mock.calls) {
      expect((call as unknown as unknown[])[2]).toBe(expand);
    }
  });

  it('rejects a malformed id with UsageError', async () => {
    const deps = buildDeps({});
    await expect(
      runListFolders(deps, { extendedProperty: ['Binary 0xZZZZ'] }),
    ).rejects.toBeInstanceOf(UsageError);
  });
});
