// src/http/extended-properties.ts
//
// Build the OData $expand clause that asks Outlook v2.0 for single-value
// extended (MAPI) properties. Used by the --extended-property flag on
// get-mail, list-mail and list-folders. See plan-003-extended-properties.md.

export class ExtendedPropertyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ExtendedPropertyError';
  }
}

/** Single-value MAPI property types accepted by the v2.0 PropertyId grammar. */
const TYPES =
  'Binary|Boolean|CLSID|Currency|Double|Float|Integer|Long|Object|Short|String|SystemTime';

/**
 * The three PropertyId forms Outlook v2.0 accepts:
 *   - tagged:     `Binary 0x348A`
 *   - named (id): `String {00020329-0000-0000-C000-000000000046} Id 0x8005`
 *   - named:      `String {00020329-0000-0000-C000-000000000046} Name Keywords`
 * Names are restricted to characters that cannot break out of the quoted
 * OData literal, so no escaping is ever needed.
 */
const GUID = '\\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\\}';
const PROPERTY_ID_RE = new RegExp(
  `^(?:${TYPES}) (?:0x[0-9A-Fa-f]{4}|${GUID} (?:Id 0x[0-9A-Fa-f]{1,8}|Name [A-Za-z0-9_.:-]+))$`,
);

/**
 * Normalise the raw flag values: each value may itself be a comma-separated
 * list, entries are trimmed, blanks dropped, duplicates removed (first
 * occurrence wins). Throws ExtendedPropertyError on a malformed id.
 */
export function parseExtendedPropertyIds(raw: readonly string[] | undefined): string[] {
  if (raw === undefined) return [];
  const out: string[] = [];
  for (const value of raw) {
    for (const part of value.split(',')) {
      const id = part.trim().replace(/\s+/g, ' ');
      if (id.length === 0) continue;
      if (!PROPERTY_ID_RE.test(id)) {
        throw new ExtendedPropertyError(
          `--extended-property must look like "Binary 0x348A" or ` +
            `"String {guid} Name X" / "String {guid} Id 0x8005", got "${id}"`,
        );
      }
      if (!out.includes(id)) out.push(id);
    }
  }
  return out;
}

/**
 * `SingleValueExtendedProperties($filter=PropertyId eq 'A' or PropertyId eq 'B')`,
 * or `''` when no ids are requested. Ids must already be validated.
 */
export function buildExtendedPropertiesExpand(ids: readonly string[]): string {
  if (ids.length === 0) return '';
  const clauses = ids.map((id) => `PropertyId eq '${id}'`).join(' or ');
  return `SingleValueExtendedProperties($filter=${clauses})`;
}
