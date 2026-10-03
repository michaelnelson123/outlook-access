# Plan 003 — Extended (MAPI) properties on get-mail, list-mail and list-folders

Plan date: 2026-09-25
Inputs consumed:

1. `docs/design/project-design.md` (§2.13.3 list-mail, §2.13.4 get-mail, §10 folders)
2. `docs/design/project-functions.MD` (FR-003, FR-004, FR-008)
3. `src/http/filter-builder.ts` (structural template for the new builder)
4. A read-only probe against a live Exchange Online mailbox (§5)

Plan 003 is **strictly additive**. Every existing invocation behaves exactly as
before: the new flag is optional, and when it is absent no `$expand` is sent.

---

## 1. Goal

Let a caller read single-value MAPI properties that the Outlook REST v2.0 schema
does not surface as named fields. The motivating case is **PidTagLastActiveParentFid**
(`Binary 0x348A`): Exchange stamps it on a message when it is moved, so a message
in Deleted Items still records the folder it was deleted _from_. Matching that
value against each folder's entry id (`Binary 0x0FFF`) tells a caller "this was
moved to Deleted Items out of folder X", which no named field can.

The flag is generic rather than hard-coded to 0x348A: any tagged or named
single-value property can be requested.

## 2. Surface

One repeatable flag on three commands:

```
outlook-cli get-mail <id> --extended-property "Binary 0x348A"
outlook-cli list-mail --folder DeletedItems --extended-property "Binary 0x348A"
outlook-cli list-folders --extended-property "Binary 0x0FFF"
```

- Repeat the flag or comma-separate values to request several properties.
- Accepted id forms (the v2.0 `PropertyId` grammar):
  - tagged — `<Type> 0x<4 hex>`, e.g. `Binary 0x348A`
  - named by id — `<Type> {<guid>} Id 0x<hex>`
  - named by name — `<Type> {<guid>} Name <name>`
  - `<Type>` is one of `Binary Boolean CLSID Currency Double Float Integer Long
Object Short String SystemTime`.
- A malformed id is a `UsageError` (exit 2) raised before any HTTP call.
- Output: each message / folder gains `SingleValueExtendedProperties:
[{ PropertyId, Value }]`, exactly as Outlook returns it. A property the item
  does not carry is simply absent from the array. `Value` is a string (base64
  for `Binary`).
- `list-mail --just-count` ignores the flag, as it ignores `--select`.

## 3. Design

- **New module `src/http/extended-properties.ts`** — `parseExtendedPropertyIds`
  (split, trim, validate, de-duplicate) and `buildExtendedPropertiesExpand`
  (`SingleValueExtendedProperties($filter=PropertyId eq 'A' or PropertyId eq 'B')`).
  Mirrors `filter-builder.ts`: a pure module with its own error class, which the
  commands translate into `UsageError`.
- **Validation is the injection guard.** The regex admits no quote, paren or
  comma, so an id can never break out of the quoted OData literal. No escaping
  code is needed.
- **Client (`src/http/outlook-client.ts`)** — `ListMessagesInFolderOptions`
  gains `expand?`, serialised as `$expand` by `buildMessagesQuery` (so both
  `listMessagesInFolder` and `listMessagesInFolderAll` carry it; `nextLink`
  pages inherit it from the server). `listFolders(parentId, top?, expand?)`
  gains an optional third argument. `get-mail` passes `$expand` through the
  existing generic `get<T>(path, query)`.
- **Types (`src/http/types.ts`)** — new `SingleValueExtendedProperty`;
  `MessageSummary` and `FolderSummary` gain an optional
  `SingleValueExtendedProperties` array.
- **CLI (`src/cli.ts`)** — `--extended-property <id>` with a small
  `collectRepeatable` parser. A variadic option (`<id...>`) was rejected: it
  would swallow `get-mail`'s positional `<id>`.
- **Configuration** — none. The flag has no config-file or env counterpart,
  so the no-fallback rule is untouched.

## 4. Tests (`test_scripts/extended-properties.spec.ts`)

- Parser: each accepted form, comma-splitting, whitespace collapse, de-dup,
  blanks; rejection of bare tags, unknown types, short tags, quote injection,
  bad GUIDs, quotes in names.
- Builder: empty, one, several (`or`-joined).
- Client: `$expand` on the wire for `listMessagesInFolder` and `listFolders`,
  and its absence when not requested (global `fetch` stubbed).
- Commands: `get-mail`, `list-mail` (single page and `--all`) and
  `list-folders` (recursive walk: every page request carries it) forward the
  clause; each rejects a malformed id with `UsageError` before HTTP.

## 5. Live verification (read-only, 2026-09-25)

Against an Exchange Online mailbox:

- `list-mail --folder DeletedItems -n 3 --extended-property "Binary 0x348A"` —
  3 of 3 messages carried the property (22-byte values).
- `list-folders --extended-property "Binary 0x0FFF"` — 26 of 26 folders
  returned an entry id; every 0x348A value was found verbatim inside exactly
  one folder's entry id.
- `get-mail <id> --extended-property "Binary 0x348A"` — property returned.

**Observed quirk:** Outlook echoes `PropertyId` with the hex lower-cased
(`Binary 0x348a`) regardless of the case requested. Callers should compare
property ids case-insensitively.

## 6. Out of scope

- Multi-value extended properties (`MultiValueExtendedProperties`).
- Writing extended properties.
- Decoding values (e.g. mapping 0x348A to a folder). That is the caller's
  job; the CLI returns what Outlook returns.
