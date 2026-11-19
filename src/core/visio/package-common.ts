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
	// A 10,000-shape page at 125 nodes a shape (the 90th percentile of the pages measured). At
	// this limit a page of plain shapes took about 6.5 s and 600 MB of V8 heap to load, the
	// densest XML 3 to 5 s and 660 MB: inside `maxRuntimeMs`, which is checked between steps
	// (DOM construction of one part is not interrupted; the viewer's worker has a hard timeout).
	maxXmlNodes: 1_250_000,
	maxTotalXmlNodes: 2_000_000,
	maxRuntimeMs: 10_000,
};
export function fail(code: string, message: string): never {
	throw new VisioPackageError(code, message);
}
export const utf8 = new TextDecoder('utf-8', { fatal: true });
export { crc32 } from '../digest/crc32';
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
