/** Limits apply to declared ZIP sizes and actual streamed output, independently. */
export interface VisioPackageLimits {
	maxInputBytes: number;
	maxEntries: number;
	maxEntryBytes: number;
	maxTotalBytes: number;
	maxCompressionRatio: number;
	maxXmlChars: number;
	maxXmlDepth: number;
	maxXmlNodes: number;
	maxTotalXmlNodes: number;
	/** Elapsed time from open(), including all subsequent reads. */
	maxRuntimeMs: number;
}

export class VisioPackageError extends Error {
	constructor(
		public readonly code: string,
		message: string,
	) {
		super(message);
		this.name = 'VisioPackageError';
	}
}

export interface VisioPackageRelationship {
	id: string;
	type: string;
	target: string;
	mode: 'Internal' | 'External';
}

export const DEFAULTS: VisioPackageLimits = {
	maxInputBytes: 32 * 1024 * 1024,
	maxEntries: 2048,
	maxEntryBytes: 16 * 1024 * 1024,
	maxTotalBytes: 128 * 1024 * 1024,
	maxCompressionRatio: 200,
	maxXmlChars: 16 * 1024 * 1024,
	maxXmlDepth: 128,
	maxXmlNodes: 250_000,
	maxTotalXmlNodes: 1_000_000,
	maxRuntimeMs: 10_000,
};
export function fail(code: string, message: string): never {
	throw new VisioPackageError(code, message);
}
export const utf8 = new TextDecoder('utf-8', { fatal: true });
const crcTable = Uint32Array.from({ length: 256 }, (_, value) => {
	for (let i = 0; i < 8; i++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
	return value >>> 0;
});
export function crc32(bytes: Uint8Array, crc = 0xffffffff): number {
	for (const byte of bytes) crc = (crcTable[(crc ^ byte) & 255] ?? 0) ^ (crc >>> 8);
	return crc >>> 0;
}
export function decodePath(path: string): string {
	try {
		return decodeURIComponent(path);
	} catch {
		return fail('UNSAFE_PATH', 'Invalid percent encoding in package path');
	}
}
/** Do not let ZIP name normalization or URI decoding turn an unsafe name into a safe one. */
export function safePath(path: string, directory = false): string {
	const decoded = decodePath(path);
	if (/%(?:2f|5c)/i.test(path)) fail('UNSAFE_PATH', 'Encoded path separators are unsupported');
	if (/[\\:#?\u0000-\u001f\u007f]/.test(path) || /[\\:#?%\u0000-\u001f\u007f]/.test(decoded))
		fail('UNSAFE_PATH', `Unsafe package path: ${path}`);
	for (const candidate of [path, decoded]) {
		const name = directory && candidate.endsWith('/') ? candidate.slice(0, -1) : candidate;
		if (!name || name.split('/').some((part) => !part || part === '.' || part === '..'))
			fail('UNSAFE_PATH', `Unsafe package path: ${path}`);
	}
	return path;
}
export interface Entry {
	name: string;
	compressed: number;
	size: number;
	crc: number;
	directory: boolean;
}
