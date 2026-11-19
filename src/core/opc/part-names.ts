/**
 * OPC part names (ECMA-376 Part 2, 6.2.2) and the ZIP item names that store them (7.3).
 *
 * A ZIP item name is a part name without its leading slash, with non-ASCII characters
 * percent-encoded. ASCII percent-encoding is part of the name: `a%20b.xml` is a part called
 * `/a%20b.xml`, never `/a b.xml`.
 */

/** The ZIP item holding the Media Types stream. It is not a part. */
export const MEDIA_TYPES_ITEM = '[Content_Types].xml';

/** Maps A-Z only. `toLowerCase` also folds non-ASCII letters, which OPC keeps distinct. */
export const asciiLowerCase = (value: string): string =>
	value.replace(/[A-Z]+/g, (letters) => letters.toLowerCase());

const utf8 = new TextDecoder('utf-8', { fatal: true });

/** Un-percent-encodes non-ASCII characters; octets that are not UTF-8 stay encoded. */
function decodeNonAscii(name: string): string {
	return name.replace(/(?:%[89a-f][0-9a-f])+/gi, (encoded) => {
		const octets = encoded.slice(1).split('%');
		try {
			return utf8.decode(Uint8Array.from(octets, (octet) => Number.parseInt(octet, 16)));
		} catch {
			return encoded;
		}
	});
}

const withoutLeadingSlash = (name: string): string => (name.startsWith('/') ? name.slice(1) : name);

/**
 * The form two names share exactly when OPC treats them as the same part: ASCII case-insensitive,
 * after mapping a ZIP item name to its part name. Accepts either form, with or without the slash.
 */
export function partNameKey(name: string): string {
	return asciiLowerCase(decodeNonAscii(withoutLeadingSlash(name)));
}

/** RFC 3987 `isegment-nz`: `ipchar` with `ucschar` spelled out. */
const SEGMENT =
	/^(?:[A-Za-z0-9\-._~!$&'()*+,;=:@\u{A0}-\u{D7FF}\u{F900}-\u{FDCF}\u{FDF0}-\u{FFEF}\u{10000}-\u{1FFFD}\u{20000}-\u{2FFFD}\u{30000}-\u{3FFFD}\u{40000}-\u{4FFFD}\u{50000}-\u{5FFFD}\u{60000}-\u{6FFFD}\u{70000}-\u{7FFFD}\u{80000}-\u{8FFFD}\u{90000}-\u{9FFFD}\u{A0000}-\u{AFFFD}\u{B0000}-\u{BFFFD}\u{C0000}-\u{CFFFD}\u{D0000}-\u{DFFFD}\u{E1000}-\u{EFFFD}]|%[0-9A-Fa-f]{2})+$/u;

const BACKSLASH = 0x5c;

/** A part name may not percent-encode a slash, a backslash or an unreserved ASCII character. */
function hasForbiddenEscape(segment: string): boolean {
	for (const [, hex] of segment.matchAll(/%([0-9A-Fa-f]{2})/g)) {
		const code = Number.parseInt(hex!, 16);
		if (code === BACKSLASH || /[A-Za-z0-9\-._~/]/.test(String.fromCharCode(code))) return true;
	}
	return false;
}

/** Whether `name` (with or without the leading slash) satisfies the part name grammar. */
export function isPartName(name: string): boolean {
	return decodeNonAscii(withoutLeadingSlash(name))
		.split('/')
		.every(
			(segment) => SEGMENT.test(segment) && !segment.endsWith('.') && !hasForbiddenEscape(segment),
		);
}
