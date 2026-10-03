// Base64 for password hash attributes, without `atob` / `btoa` (whose error handling and
// availability differ between runtimes).

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
const VALUES = new Map([...ALPHABET].map((char, index) => [char, index]));

/** Encodes `bytes` as padded base64. */
export function encodeBase64(bytes: Uint8Array): string {
	let out = '';
	for (let i = 0; i < bytes.length; i += 3) {
		const a = bytes[i]!;
		const b = bytes[i + 1];
		const c = bytes[i + 2];
		const n = (a << 16) | ((b ?? 0) << 8) | (c ?? 0);
		out += ALPHABET[n >> 18]! + ALPHABET[(n >> 12) & 63]!;
		out += b === undefined ? '=' : ALPHABET[(n >> 6) & 63]!;
		out += c === undefined ? '=' : ALPHABET[n & 63]!;
	}
	return out;
}

/**
 * Decodes base64, ignoring white space and tolerating missing padding. Returns `undefined` for
 * anything malformed (characters outside the alphabet, misplaced padding, an impossible length).
 */
export function decodeBase64(text: string): Uint8Array | undefined {
	const compact = text.replace(/\s+/g, '');
	const body = compact.replace(/={1,2}$/, '');
	if (body.length % 4 === 1) return undefined;
	if (compact.length !== body.length && compact.length % 4 !== 0) return undefined;
	const out = new Uint8Array(Math.floor((body.length * 3) / 4));
	let bits = 0;
	let buffer = 0;
	let index = 0;
	for (const char of body) {
		const value = VALUES.get(char);
		if (value === undefined) return undefined;
		buffer = ((buffer << 6) | value) & 0xffffff;
		bits += 6;
		if (bits >= 8) {
			bits -= 8;
			out[index++] = (buffer >> bits) & 0xff;
		}
	}
	return out;
}
