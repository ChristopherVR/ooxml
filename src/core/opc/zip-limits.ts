import JSZip from 'jszip';

/** Default limits for a nested or unspecified package. */
export const DEFAULT_ZIP_LIMITS = {
	maxInputBytes: 100 * 1024 * 1024,
	maxEntries: 10_000,
	maxEntryBytes: 100 * 1024 * 1024,
	maxTotalBytes: 100 * 1024 * 1024,
} as const;

/**
 * The check a package failed: one of the {@link ZipLimits} sizes, `declaredSize` (an entry
 * inflates past the size its header declares) or `partName` (an unsafe part name).
 */
export type ZipLimitKind =
	| 'maxInputBytes'
	| 'maxEntries'
	| 'maxEntryBytes'
	| 'maxTotalBytes'
	| 'declaredSize'
	| 'partName';

/** Size and hygiene limits applied while opening and reading a zip package. */
export interface ZipLimits {
	/** Largest accepted package (compressed bytes). */
	maxInputBytes?: number;
	/** Largest number of zip entries. */
	maxEntries?: number;
	/** Largest inflated size of one entry. */
	maxEntryBytes?: number;
	/** Largest total inflated size of every entry read or declared. */
	maxTotalBytes?: number;
	/** Format name used in error messages, such as `DOCX`. */
	label?: string;
	/** Builds the error thrown on a violation; defaults to {@link ZipLimitError}. */
	createError?: (message: string, details: { limit: number; kind: ZipLimitKind }) => Error;
}

/** Thrown when a package breaks a size, entry-count or part-name limit. */
export class ZipLimitError extends Error {
	readonly code = 'ZIP_BOMB';
	readonly limit: number;
	/** The check that failed (absent when the error was built without one). */
	readonly kind: ZipLimitKind | undefined;
	constructor(message: string, details: { limit: number; kind?: ZipLimitKind }) {
		super(message);
		this.name = 'ZipLimitError';
		this.limit = details.limit;
		this.kind = details.kind;
	}
}

const mib = (bytes: number): number => Math.round(bytes / 1048576);

/** True when a zip entry name is a safe, relative part name (no traversal, drive or backslash). */
export function isSafePartName(name: string): boolean {
	if (name.includes('\0') || name.includes('\\')) return false;
	if (/^[A-Za-z]:/u.test(name)) return false;
	const body = name.startsWith('/') ? name.slice(1) : name;
	if (body.startsWith('/')) return false;
	return !body.split('/').some((segment) => segment === '..');
}

/** The declared (header) uncompressed size of an entry; attacker-controlled. */
export function declaredEntrySize(entry: JSZip.JSZipObject): number {
	const data = (entry as unknown as { _data?: { uncompressedSize?: unknown } })._data;
	const size = Number(data?.uncompressedSize ?? 0);
	return Number.isFinite(size) && size > 0 ? size : 0;
}

interface InternalStream {
	internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
}

interface Resolved {
	maxEntries: number;
	maxEntryBytes: number;
	maxTotalBytes: number;
	label: string;
	fail: (message: string, limit: number, kind: ZipLimitKind) => Error;
}

function resolveLimits(limits: ZipLimits): Resolved {
	const create = limits.createError;
	return {
		maxEntries: limits.maxEntries ?? DEFAULT_ZIP_LIMITS.maxEntries,
		maxEntryBytes: limits.maxEntryBytes ?? DEFAULT_ZIP_LIMITS.maxEntryBytes,
		maxTotalBytes: limits.maxTotalBytes ?? DEFAULT_ZIP_LIMITS.maxTotalBytes,
		label: limits.label ?? 'ZIP',
		fail: (message, limit, kind) =>
			create ? create(message, { limit, kind }) : new ZipLimitError(message, { limit, kind }),
	};
}

/**
 * Inflates one entry chunk by chunk, counting the bytes actually produced. Header sizes are
 * attacker-controlled, so the read is rejected as soon as the entry inflates past its declared
 * size or the running total passes the limit, before the rest of the entry is inflated.
 */
export function inflateEntryGuarded(
	entry: JSZip.JSZipObject,
	alreadyRead: number,
	limits: ZipLimits,
): Promise<Uint8Array> {
	const r = resolveLimits(limits);
	const declared = declaredEntrySize(entry);
	return new Promise((resolvePromise, reject) => {
		const chunks: Uint8Array[] = [];
		let size = 0;
		let done = false;
		const stream = (entry as unknown as InternalStream).internalStream('uint8array');
		const abort = (error: Error): void => {
			if (done) return;
			done = true;
			stream.pause();
			chunks.length = 0;
			reject(error);
		};
		stream.on('data', (chunk) => {
			if (done) return;
			size += chunk.byteLength;
			if (size > declared || size > r.maxEntryBytes)
				return abort(
					r.fail(
						`${r.label} part ${entry.name} inflates past its declared size`,
						r.maxEntryBytes,
						size > declared ? 'declaredSize' : 'maxEntryBytes',
					),
				);
			if (alreadyRead + size > r.maxTotalBytes)
				return abort(
					r.fail(
						`${r.label} package exceeds the ${mib(r.maxTotalBytes)} MiB uncompressed content limit`,
						r.maxTotalBytes,
						'maxTotalBytes',
					),
				);
			chunks.push(chunk);
		});
		stream.on('error', (error) => abort(error));
		stream.on('end', () => {
			if (done) return;
			done = true;
			if (chunks.length === 1 && chunks[0]) return resolvePromise(chunks[0]);
			const out = new Uint8Array(size);
			let offset = 0;
			for (const chunk of chunks) {
				out.set(chunk, offset);
				offset += chunk.byteLength;
			}
			resolvePromise(out);
		});
		stream.resume();
	});
}

function toBase64(bytes: Uint8Array): string {
	let binary = '';
	for (let i = 0; i < bytes.length; i += 0x8000)
		binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
	return btoa(binary);
}

/**
 * Validates a loaded zip (entry count, part names, declared sizes) and replaces every entry's
 * `async` so each read inflates through the streaming guard with per-entry and running-total
 * caps. Existing `zip.file(name).async(...)` call sites are therefore protected unchanged.
 */
export function guardZip(zip: JSZip, limits: ZipLimits = {}): JSZip {
	const r = resolveLimits(limits);
	const entries = Object.values(zip.files);
	if (entries.length > r.maxEntries)
		throw r.fail(
			`${r.label} package exceeds the ${r.maxEntries.toLocaleString('en-US')} part limit`,
			r.maxEntries,
			'maxEntries',
		);
	let declared = 0;
	for (const entry of entries) {
		const original = (entry as unknown as { unsafeOriginalName?: string }).unsafeOriginalName;
		if (!isSafePartName(entry.name) || (original !== undefined && !isSafePartName(original)))
			throw r.fail(`${r.label} package has an unsafe part name`, r.maxEntries, 'partName');
		if (entry.dir) continue;
		const size = declaredEntrySize(entry);
		if (size > r.maxEntryBytes)
			throw r.fail(
				`${r.label} part ${entry.name} exceeds the uncompressed part size limit`,
				r.maxEntryBytes,
				'maxEntryBytes',
			);
		declared += size;
	}
	if (declared > r.maxTotalBytes)
		throw r.fail(
			`${r.label} package exceeds the ${mib(r.maxTotalBytes)} MiB uncompressed content limit`,
			r.maxTotalBytes,
			'maxTotalBytes',
		);
	const produced = new Map<string, number>();
	let total = 0;
	for (const entry of entries) {
		if (entry.dir) continue;
		const inflateOriginal = entry.async.bind(entry) as (
			type: string,
			onUpdate?: unknown,
		) => Promise<unknown>;
		(entry as unknown as { async: unknown }).async = async (type: string, onUpdate?: unknown) => {
			const previous = produced.get(entry.name) ?? 0;
			const bytes = await inflateEntryGuarded(entry, total - previous, limits);
			if (bytes.byteLength > previous) {
				total += bytes.byteLength - previous;
				produced.set(entry.name, bytes.byteLength);
			}
			switch (type) {
				case 'uint8array':
					return bytes;
				case 'string':
				case 'text':
					return new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes);
				case 'arraybuffer':
					return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
				case 'base64':
					return toBase64(bytes);
				default:
					return inflateOriginal(type, onUpdate);
			}
		};
	}
	return zip;
}

/** Opens a zip with the input-size cap, then guards it (see {@link guardZip}). */
export async function openZip(
	input: Uint8Array | ArrayBuffer,
	limits: ZipLimits = {},
): Promise<JSZip> {
	const r = resolveLimits(limits);
	const maxInput = limits.maxInputBytes ?? DEFAULT_ZIP_LIMITS.maxInputBytes;
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (bytes.byteLength > maxInput)
		throw r.fail(
			`${r.label} package exceeds the ${mib(maxInput)} MiB compressed input limit`,
			maxInput,
			'maxInputBytes',
		);
	return guardZip(await JSZip.loadAsync(bytes), limits);
}
