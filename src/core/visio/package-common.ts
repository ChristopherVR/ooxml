import type { ZipLimitKind, ZipLimits } from '../opc/zip-limits';
import { defaultXmlHeapBytes } from './heap-limit';

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
	/**
	 * Estimated V8 heap, in bytes, one read or edit may hold: the parsed XML of every package it
	 * opens (twice over for an edit, which may copy what it reads) and, for a read, the model.
	 * Each part is costed while it is scanned, before its DOM is built: 800 bytes an element, 240
	 * an attribute, 140 any other node and 1 a character, as measured for the shared XML parser.
	 * By default, the host's heap limit less 256 MiB, divided by 1.6 (the peak heap measured per
	 * byte charged), or 480 MiB where the runtime reports no limit.
	 */
	maxXmlHeapBytes: number;
	/**
	 * The same estimate per byte of the input file, once it passes 64 MiB, so a small file cannot
	 * expand into a large heap.
	 */
	maxXmlHeapPerInputByte: number;
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
	// Bounds the DOM construction of one part, which `maxRuntimeMs` cannot interrupt: a
	// 10,000-shape page at 150 nodes a shape (Visio saves 128 for a Basic Shapes rectangle).
	maxXmlNodes: 1_500_000,
	maxTotalXmlNodes: 2_000_000,
	// Memory is bounded by the heap estimate, not the node counts, and the estimate follows the
	// host's heap so no drawing can exhaust it (README, Bounded processing).
	maxXmlHeapBytes: defaultXmlHeapBytes(),
	// Drawings Visio saves estimate 90 to 3,100 bytes per input byte, generated pages of identical
	// shapes about 2,100; this refuses the densest XML a 50 KB file can hold before it is parsed.
	maxXmlHeapPerInputByte: 4096,
	maxRuntimeMs: 10_000,
};
export function fail(code: string, message: string): never {
	throw new VisioPackageError(code, message);
}

const ZIP_LIMIT_CODES: Record<ZipLimitKind, string> = {
	maxInputBytes: 'LIMIT_INPUT',
	maxEntries: 'LIMIT_ENTRIES',
	maxEntryBytes: 'LIMIT_ENTRY',
	maxTotalBytes: 'LIMIT_TOTAL',
	declaredSize: 'ZIP_MISMATCH',
	partName: 'UNSAFE_PATH',
};

/** The zip guard settings for `limits`; a violation is a {@link VisioPackageError}. */
export function visioZipLimits(limits: VisioPackageLimits): ZipLimits {
	return {
		label: 'Visio',
		maxInputBytes: limits.maxInputBytes,
		maxEntries: limits.maxEntries,
		maxEntryBytes: limits.maxEntryBytes,
		maxTotalBytes: limits.maxTotalBytes,
		createError: (message, { kind }) => new VisioPackageError(ZIP_LIMIT_CODES[kind], message),
	};
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
