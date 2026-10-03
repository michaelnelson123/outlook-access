// test_scripts/outlook-client-threads.spec.ts
//
// Tests for the thread + date-filter additions on `OutlookClient`:
//   - `listMessagesInFolder` now accepts `filter` (threaded into $filter)
//   - `listMessagesByConversation` (new method)
//
// Kept sibling to outlook-client-folders.spec.ts; same mocking style.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createOutlookClient } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';
import type { MessageSummary } from '../src/http/types';

const JWT_SHAPED_TOKEN = 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

function buildFakeSession(): SessionFile {
  return {
    version: 1,
    capturedAt: '2026-04-21T12:00:00.000Z',
    account: {
      upn: 'alice@contoso.com',
      puid: '1234567890',
      tenantId: 'tenant-id-abc',
    },
    bearer: {
      token: JWT_SHAPED_TOKEN,
      expiresAt: '2099-04-21T12:00:00.000Z',
      audience: 'https://outlook.office.com',
      scopes: ['Mail.Read'],
    },
    cookies: [
      {
        name: 'SessionCookie',
        value: 'outlook-cookie-value',
        domain: '.outlook.office.com',
        path: '/',
        expires: -1,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
    ],
    anchorMailbox: 'PUID:1234567890@tenant-id-abc',
  };
}

function makeResponse(init: { status: number; body: unknown }): Response {
  const headersMap = new Headers();
  const bodyText = JSON.stringify(init.body);
  return {
    status: init.status,
    ok: init.status >= 200 && init.status < 300,
    headers: headersMap,
    text: async () => bodyText,
    json: async () => JSON.parse(bodyText),
  } as unknown as Response;
}

function makeMessage(id: string, received: string): MessageSummary {
  return {
    Id: id,
    Subject: `Subject ${id}`,
    ReceivedDateTime: received,
    HasAttachments: false,
    IsRead: false,
    WebLink: `https://example.com/${id}`,
  };
}

describe('listMessagesInFolder — filter option', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('threads a raw $filter expression into the query string', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [makeMessage('m1', '2026-04-01T10:00:00Z')] } }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listMessagesInFolder('Inbox', {
      top: 10,
      filter:
        'ReceivedDateTime ge 2026-04-01T00:00:00.000Z and ReceivedDateTime lt 2026-05-01T00:00:00.000Z',
    });

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).toContain('%24filter=');
    // URL-encoded form of the literal single-quote-free filter
    expect(decodeURIComponent(url.replace(/\+/g, '%20'))).toContain(
      'ReceivedDateTime ge 2026-04-01T00:00:00.000Z and ReceivedDateTime lt 2026-05-01T00:00:00.000Z',
    );
  });

  it('omits $filter when option is not provided', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { value: [] } }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listMessagesInFolder('Inbox', { top: 5 });
    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).not.toContain('%24filter=');
  });
});

describe('listMessagesByConversation', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('builds the expected /messages URL with $filter=ConversationId eq (no $orderby — sorted client-side)', async () => {
    // Returned out-of-order to verify client-side sort applies.
    const msgs = [
      makeMessage('m2', '2026-03-01T10:00:00Z'),
      makeMessage('m1', '2026-03-01T09:00:00Z'),
    ];
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { value: msgs } }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listMessagesByConversation('CONV-ABC-123');
    // Default order: asc → m1 (09:00) before m2 (10:00).
    expect(result.map((m) => m.Id)).toEqual(['m1', 'm2']);
    expect(fetchMock).toHaveBeenCalledTimes(1);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).toContain('https://outlook.office.com/api/v2.0/me/messages?');
    const decoded = decodeURIComponent(url.replace(/\+/g, '%20'));
    expect(decoded).toContain("ConversationId eq 'CONV-ABC-123'");
    // Server-side $orderby intentionally omitted (Outlook rejects this combo
    // as InefficientFilter — see fork CHANGELOG 1.2.0). Sorting is client-side.
    expect(decoded).not.toContain('$orderby');
  });

  it('honors custom orderBy (client-side desc sort) and forwards select/top', async () => {
    // Out-of-order; orderBy: desc should reverse them.
    const msgs = [
      makeMessage('a', '2026-03-01T09:00:00Z'),
      makeMessage('b', '2026-03-01T10:00:00Z'),
    ];
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { value: msgs } }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listMessagesByConversation('CID', {
      orderBy: 'ReceivedDateTime desc',
      select: ['Id', 'Subject', 'Body'],
      top: 50,
    });
    // desc → b (10:00) before a (09:00).
    expect(result.map((m) => m.Id)).toEqual(['b', 'a']);

    const decoded = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].toString().replace(/\+/g, '%20'),
    );
    // Client-side sort, so server-side $orderby is omitted.
    expect(decoded).not.toContain('$orderby');
    expect(decoded).toContain('Id,Subject,Body');
    expect(decoded).toContain('$top=50');
  });

  it("escapes single quotes inside the conversation id (OData ' → '')", async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { value: [] } }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listMessagesByConversation("weird'id");
    const decoded = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].toString().replace(/\+/g, '%20'),
    );
    expect(decoded).toContain("ConversationId eq 'weird''id'");
  });

  it('throws when conversationId is empty', async () => {
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
    await expect(client.listMessagesByConversation('')).rejects.toThrow(/non-empty conversationId/);
  });
});

describe('listMessagesByConversation — paging (server default page of 10 truncated threads)', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function newClient() {
    return createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
  }

  it('follows @odata.nextLink, stitches both pages and sorts the union client-side', async () => {
    const nextLink =
      "https://outlook.office.com/api/v2.0/me/messages?$filter=ConversationId+eq+'CID'&$top=250&$skip=250";
    // Newest message deliberately on page 2, as the server returned it live.
    fetchMock
      .mockResolvedValueOnce(
        makeResponse({
          status: 200,
          body: {
            value: [
              makeMessage('p1-b', '2026-03-02T10:00:00Z'),
              makeMessage('p1-a', '2026-03-01T10:00:00Z'),
            ],
            '@odata.nextLink': nextLink,
          },
        }),
      )
      .mockResolvedValueOnce(
        makeResponse({
          status: 200,
          body: {
            value: [
              makeMessage('p2-newest', '2026-03-09T10:00:00Z'),
              makeMessage('p2-oldest', '2026-02-01T10:00:00Z'),
            ],
          },
        }),
      );

    const result = await newClient().listMessagesByConversation('CID');

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(result.map((m) => m.Id)).toEqual(['p2-oldest', 'p1-a', 'p1-b', 'p2-newest']);

    const first = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].replace(/\+/g, '%20'),
    );
    // Without an explicit top the listAll page size applies, not the server's 10.
    expect(first).toContain('$top=250');
    expect(first).not.toContain('$orderby');
    // The nextLink is followed verbatim.
    expect((fetchMock.mock.calls[1] as [string, unknown])[0]).toBe(nextLink);
  });

  it('honours an explicit top: sent as $top and caps the result across pages', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: {
          value: [
            makeMessage('a', '2026-03-01T10:00:00Z'),
            makeMessage('b', '2026-03-02T10:00:00Z'),
          ],
          '@odata.nextLink': 'https://outlook.office.com/api/v2.0/me/messages?$skip=2',
        },
      }),
    );

    const result = await newClient().listMessagesByConversation('CID', { top: 2 });

    // Cap reached on page 1, so the nextLink is not fetched.
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.map((m) => m.Id)).toEqual(['a', 'b']);
    const decoded = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].replace(/\+/g, '%20'),
    );
    expect(decoded).toContain('$top=2');
  });

  it('refuses an off-host @odata.nextLink with UPSTREAM_PAGINATION_LIMIT', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: {
          value: [makeMessage('a', '2026-03-01T10:00:00Z')],
          '@odata.nextLink': 'https://evil.example.com/api/v2.0/me/messages?$skip=1',
        },
      }),
    );

    await expect(newClient().listMessagesByConversation('CID')).rejects.toMatchObject({
      code: 'UPSTREAM_PAGINATION_LIMIT',
    });
    // The off-host URL is never fetched.
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe('Prefer: outlook.body-content-type="text"', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function newClient() {
    return createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
  }

  function headersOf(call: number): Record<string, string> {
    return (fetchMock.mock.calls[call] as [string, RequestInit])[1].headers as Record<
      string,
      string
    >;
  }

  it('get() sends the Prefer header when bodyContentType is text', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { Id: 'm1' } }));
    await newClient().get('/api/v2.0/me/messages/m1', undefined, { bodyContentType: 'text' });
    expect(headersOf(0).Prefer).toBe('outlook.body-content-type="text"');
  });

  it('get() sends no Prefer header by default or for html', async () => {
    fetchMock
      .mockResolvedValueOnce(makeResponse({ status: 200, body: { Id: 'm1' } }))
      .mockResolvedValueOnce(makeResponse({ status: 200, body: { Id: 'm1' } }));
    const client = newClient();
    await client.get('/api/v2.0/me/messages/m1');
    await client.get('/api/v2.0/me/messages/m1', undefined, { bodyContentType: 'html' });
    expect(headersOf(0).Prefer).toBeUndefined();
    expect(headersOf(1).Prefer).toBeUndefined();
  });

  it('listMessagesByConversation sends the Prefer header on every page for text', async () => {
    fetchMock
      .mockResolvedValueOnce(
        makeResponse({
          status: 200,
          body: {
            value: [makeMessage('a', '2026-03-01T10:00:00Z')],
            '@odata.nextLink': 'https://outlook.office.com/api/v2.0/me/messages?$skip=1',
          },
        }),
      )
      .mockResolvedValueOnce(
        makeResponse({ status: 200, body: { value: [makeMessage('b', '2026-03-02T10:00:00Z')] } }),
      );
    await newClient().listMessagesByConversation('CID', { bodyContentType: 'text' });
    expect(headersOf(0).Prefer).toBe('outlook.body-content-type="text"');
    expect(headersOf(1).Prefer).toBe('outlook.body-content-type="text"');
  });

  it('listMessagesByConversation sends no Prefer header without bodyContentType', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { value: [] } }));
    await newClient().listMessagesByConversation('CID');
    expect(headersOf(0).Prefer).toBeUndefined();
  });
});

describe('updateMessage — IsRead / Flag PATCH bodies (mark-mail)', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it.each([
    [{ IsRead: true }],
    [{ IsRead: false }],
    [{ Flag: { FlagStatus: 'Flagged' as const } }],
    [{ Flag: { FlagStatus: 'NotFlagged' as const } }],
    [{ Flag: { FlagStatus: 'Complete' as const } }],
  ])('PATCHes /me/messages/{id} with %j', async (patch) => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { Id: 'm1' } }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
    await client.updateMessage('m1', patch);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe('https://outlook.office.com/api/v2.0/me/messages/m1');
    expect(init.method).toBe('PATCH');
    expect(JSON.parse(init.body as string)).toEqual(patch);
  });
});

describe('countMessagesInFolder', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('sends $count=true&$top=1&$select=Id and returns @odata.count as exact:true', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { '@odata.count': 4273, value: [makeMessage('m1', '2026-04-01T10:00:00Z')] },
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.countMessagesInFolder('Inbox');
    expect(result.count).toBe(4273);
    expect(result.exact).toBe(true);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    const decoded = decodeURIComponent(url.replace(/\+/g, '%20'));
    expect(decoded).toContain('$count=true');
    expect(decoded).toContain('$top=1');
    expect(decoded).toContain('$select=Id');
    expect(url).toContain('/MailFolders/Inbox/messages');
  });

  it('threads filter into the request', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { '@odata.count': 12, value: [] },
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.countMessagesInFolder('AAMk-raw-id', {
      filter: 'ReceivedDateTime ge 2026-04-01T00:00:00Z',
    });
    const decoded = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].toString().replace(/\+/g, '%20'),
    );
    expect(decoded).toContain('ReceivedDateTime ge 2026-04-01T00:00:00Z');
  });

  it('falls back to value.length with exact:false when server omits @odata.count', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { value: [makeMessage('m1', '2026-04-01T10:00:00Z')] }, // no @odata.count
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.countMessagesInFolder('Inbox');
    expect(result.count).toBe(1);
    expect(result.exact).toBe(false);
  });

  it('throws when folderId is empty', async () => {
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
    await expect(client.countMessagesInFolder('')).rejects.toThrow(/non-empty folderId/);
  });
});
