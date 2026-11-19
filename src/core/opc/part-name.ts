const TOKEN_SAFE = /^[A-Za-z0-9\-._~/]$/u;
const utf8 = new TextEncoder();

/**
 * A part name as text that is safe to build identifiers from.
 *
 * Part names come from the ZIP directory and relationship targets of an untrusted
 * package, and a reader cannot rely on them following the OPC grammar. Ordinary names
 * pass through unchanged, so identifiers derived from them stay stable. Every other
 * character becomes percent-encoded UTF-8, `%` included, so two different names can
 * never share a token and the result holds nothing markup or CSS treats as syntax.
 */
export function partNameToken(partName: string): string {
	let token = '';
	for (const char of partName) {
		if (TOKEN_SAFE.test(char)) {
			token += char;
			continue;
		}
		for (const byte of utf8.encode(char)) {
			token += `%${byte.toString(16).toUpperCase().padStart(2, '0')}`;
		}
	}
	return token;
}
