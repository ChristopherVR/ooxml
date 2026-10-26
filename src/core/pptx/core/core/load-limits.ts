/**
 * Host-configurable limits on the work one `PptxHandler.load()` call may do.
 *
 * The ZIP limits (`maxUncompressedBytes`, `maxEntries`) are enforced by the
 * zip guard and keep raising {@link ZipBombError}; the others raise
 * {@link PptxLoadLimitError}. {@link getPptxLoadLimitViolation} reads either.
 *
 * @module pptx-core/load-limits
 */
import type { ZipLimitKind } from '../../../opc/zip-limits';
import type { PptxElement } from '../types';
import { isTemplateElementId } from '../utils/group-ops';
import { DEFAULT_MAX_UNCOMPRESSED_BYTES, MAX_ZIP_ENTRY_COUNT, ZipBombError } from './types';
import type { PptxHandlerLoadOptions } from './types';

/** Limits on one presentation load; each field left out keeps its default. */
export interface PptxLoadLimits {
	/**
	 * Largest accepted input, in bytes, as passed to `load()`; for
	 * `loadFromJson()`, the UTF-8 size of the text (default: no limit).
	 */
	maxInputBytes?: number;
	/** Largest total inflated size of the package, in bytes (default 500 MiB). */
	maxUncompressedBytes?: number;
	/** Largest number of ZIP entries (default and ceiling 65,536). */
	maxEntries?: number;
	/** Largest number of slides listed in `ppt/presentation.xml` (default: no limit). */
	maxSlides?: number;
	/**
	 * Largest number of elements authored on the slides, summed over the deck
	 * (shapes, pictures, connectors, tables, charts, groups and every element
	 * inside a group, at any depth). Elements inherited from the layout or
	 * master (ids prefixed `layout-` or `master-`) are not counted, in a
	 * `.pptx` package or a `pptx-viewer-json` document alike (default: no limit).
	 */
	maxShapes?: number;
}

/** The name of one load limit. */
export type PptxLoadLimitName = keyof PptxLoadLimits;

/** The limits a load applies when the caller sets none. `Infinity` means no limit. */
export const DEFAULT_PPTX_LOAD_LIMITS: Readonly<Required<PptxLoadLimits>> = Object.freeze({
	maxInputBytes: Number.POSITIVE_INFINITY,
	maxUncompressedBytes: DEFAULT_MAX_UNCOMPRESSED_BYTES,
	maxEntries: MAX_ZIP_ENTRY_COUNT,
	maxSlides: Number.POSITIVE_INFINITY,
	maxShapes: Number.POSITIVE_INFINITY,
});

const LIMIT_NAMES = Object.keys(DEFAULT_PPTX_LOAD_LIMITS) as PptxLoadLimitName[];

/**
 * Thrown by `load()` when the input, the slide count or the shape count passes
 * its limit (`maxInputBytes`, `maxSlides`, `maxShapes`). The ZIP limits raise
 * `ZipBombError` as they always have. Each entry point bundles its own copy of
 * this class, so test `code` (or use {@link getPptxLoadLimitViolation}) rather
 * than `instanceof` when the error may come from another entry point.
 */
export class PptxLoadLimitError extends Error {
	public readonly code = 'PPTX_LOAD_LIMIT';

	/** The limit that was exceeded. */
	public readonly limit: PptxLoadLimitName;

	/** The value of that limit. */
	public readonly maximum: number;

	/**
	 * The observed value: exact for `maxSlides` and for `maxInputBytes` on
	 * `load()`; for `maxShapes`, and for `maxInputBytes` on `loadFromJson()`
	 * (which stops counting once past the limit), a lower bound.
	 */
	public readonly actual: number;

	constructor(message: string, limit: PptxLoadLimitName, maximum: number, actual: number) {
		super(message);
		this.name = 'PptxLoadLimitError';
		this.limit = limit;
		this.maximum = maximum;
		this.actual = actual;
	}
}

/** A load that failed on one of the {@link PptxLoadLimits}. */
export interface PptxLoadLimitViolation {
	limit: PptxLoadLimitName;
	maximum: number;
	/** The observed value, when known. For streamed sizes and shapes it is a lower bound. */
	actual: number | undefined;
}

function isLimitName(value: unknown): value is PptxLoadLimitName {
	return typeof value === 'string' && (LIMIT_NAMES as string[]).includes(value);
}

/**
 * The limit a failed load broke, from a {@link PptxLoadLimitError} or a
 * `ZipBombError` raised by a size or entry-count limit; `undefined` for any
 * other error (including a forged size or an unsafe part name, which no limit
 * lets through). Matches by `code`, so it works across bundles.
 */
export function getPptxLoadLimitViolation(error: unknown): PptxLoadLimitViolation | undefined {
	if (typeof error !== 'object' || error === null) {
		return undefined;
	}
	const fields = error as {
		code?: unknown;
		limit?: unknown;
		loadLimit?: unknown;
		maximum?: unknown;
		actual?: unknown;
	};
	const actual = typeof fields.actual === 'number' ? fields.actual : undefined;
	if (
		fields.code === 'PPTX_LOAD_LIMIT' &&
		isLimitName(fields.limit) &&
		typeof fields.maximum === 'number'
	) {
		return { limit: fields.limit, maximum: fields.maximum, actual };
	}
	if (
		fields.code === 'ZIP_BOMB' &&
		isLimitName(fields.loadLimit) &&
		typeof fields.limit === 'number'
	) {
		return { limit: fields.loadLimit, maximum: fields.limit, actual };
	}
	return undefined;
}

/**
 * The limits for `options`: the defaults, then the legacy top-level
 * `maxUncompressedBytes` (ignored unless positive, as before), then `limits`.
 * A value in `limits` must be a positive finite number, and `maxEntries` at
 * most 65,536, so a computed `NaN` or `Infinity` cannot switch a guard off.
 *
 * @throws RangeError for an invalid value in `options.limits`.
 */
export function resolvePptxLoadLimits(
	options: Pick<PptxHandlerLoadOptions, 'limits' | 'maxUncompressedBytes'> = {},
): Required<PptxLoadLimits> {
	const resolved: Required<PptxLoadLimits> = { ...DEFAULT_PPTX_LOAD_LIMITS };
	const legacy = options.maxUncompressedBytes;
	if (typeof legacy === 'number' && legacy > 0) {
		resolved.maxUncompressedBytes = legacy;
	}
	const limits = options.limits;
	if (limits === undefined || limits === null) {
		return resolved;
	}
	for (const name of LIMIT_NAMES) {
		const value = limits[name];
		if (value === undefined) {
			continue;
		}
		if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
			throw new RangeError(`limits.${name} must be a positive finite number`);
		}
		if (name === 'maxEntries' && value > MAX_ZIP_ENTRY_COUNT) {
			throw new RangeError(`limits.maxEntries must not exceed ${MAX_ZIP_ENTRY_COUNT}`);
		}
		resolved[name] = value;
	}
	return resolved;
}

/**
 * The `ZipBombError` for a zip guard violation, naming the load limit behind
 * it (`maxEntries`, or `maxUncompressedBytes` for a part or total size) and
 * the observed value. The message is the guard's own, unchanged.
 */
export function createPptxZipBombError(
	message: string,
	details: { limit: number; kind: ZipLimitKind; actual?: number | undefined },
): ZipBombError {
	const { limit, kind, actual } = details;
	if (kind === 'maxEntries') {
		return new ZipBombError(message, {
			limit,
			loadLimit: 'maxEntries',
			...(actual === undefined ? {} : { entryCount: actual, actual }),
		});
	}
	if (kind === 'maxEntryBytes' || kind === 'maxTotalBytes') {
		return new ZipBombError(message, {
			limit,
			loadLimit: 'maxUncompressedBytes',
			...(actual === undefined ? {} : { uncompressedBytes: actual, actual }),
		});
	}
	return new ZipBombError(message, { limit });
}

const count = (value: number): string => value.toLocaleString('en-US');

/** Fails when `byteLength` (the input passed to `load()`) passes `maxInputBytes`. */
export function assertPptxInputWithinLimits(
	byteLength: number,
	limits: Required<PptxLoadLimits>,
): void {
	if (byteLength > limits.maxInputBytes) {
		throw new PptxLoadLimitError(
			`PPTX input is ${count(byteLength)} bytes, more than the maxInputBytes limit of ${count(limits.maxInputBytes)}`,
			'maxInputBytes',
			limits.maxInputBytes,
			byteLength,
		);
	}
}

/**
 * Fails when `text` (given to `loadFromJson()`) takes more than
 * `maxInputBytes` bytes as UTF-8, the size the same document has as bytes
 * for `load()`. The size is counted from the UTF-16 code units without
 * encoding or copying the text, and the count stops once past the limit, so
 * the work is bounded by the limit rather than by the text. `actual` is then
 * a lower bound.
 */
export function assertPptxTextWithinLimits(text: string, limits: Required<PptxLoadLimits>): void {
	const maximum = limits.maxInputBytes;
	// A code unit takes 1 to 3 bytes (a surrogate pair, two units, takes 4).
	if (text.length * 3 <= maximum) {
		return;
	}
	const bytes = text.length > maximum ? text.length : utf8ByteLengthPast(text, maximum);
	if (bytes > maximum) {
		throw new PptxLoadLimitError(
			`PPTX input is at least ${count(bytes)} bytes, more than the maxInputBytes limit of ${count(maximum)}`,
			'maxInputBytes',
			maximum,
			bytes,
		);
	}
}

/**
 * The UTF-8 size of `text` (as `TextEncoder` writes it: a lone surrogate
 * becomes the 3-byte U+FFFD), or the first running total past `maximum`.
 */
function utf8ByteLengthPast(text: string, maximum: number): number {
	let bytes = 0;
	for (let index = 0; index < text.length && bytes <= maximum; index++) {
		const unit = text.charCodeAt(index);
		if (unit < 0x80) {
			bytes += 1;
		} else if (unit < 0x800) {
			bytes += 2;
		} else if (unit >= 0xd800 && unit <= 0xdbff && isLowSurrogate(text.charCodeAt(index + 1))) {
			bytes += 4;
			index++;
		} else {
			bytes += 3;
		}
	}
	return bytes;
}

function isLowSurrogate(unit: number): boolean {
	return unit >= 0xdc00 && unit <= 0xdfff;
}

/** Fails when the deck lists more than `maxSlides` slides. */
export function assertPptxSlideCountWithinLimits(
	slideCount: number,
	limits: Required<PptxLoadLimits>,
): void {
	if (slideCount > limits.maxSlides) {
		throw new PptxLoadLimitError(
			`PPTX presentation has ${count(slideCount)} slides, more than the maxSlides limit of ${count(limits.maxSlides)}`,
			'maxSlides',
			limits.maxSlides,
			slideCount,
		);
	}
}

/**
 * Adds the elements one slide authored to `counted` and fails as soon as the
 * total passes `maxShapes`. Returns the new total.
 *
 * The one counting rule for both load paths (a slide's own parse in a
 * `.pptx` package, a slide's `elements` in a `pptx-viewer-json` document):
 * every element counts, group children included at any depth, except an
 * element inherited from the layout or master and everything inside it. The
 * package loader marks those with a `layout-` / `master-` id prefix, and the
 * JSON document keeps the ids, so the same deck counts the same either way.
 *
 * The walk is iterative with no depth cap, so it visits each element of the
 * slide once whatever the nesting: its work and memory are bounded by the
 * elements in the input. Entries that are not objects (possible in a JSON
 * document) are skipped.
 */
export function countPptxShapesWithinLimits(
	counted: number,
	elements: readonly PptxElement[] | readonly unknown[],
	limits: Required<PptxLoadLimits>,
): number {
	const maximum = limits.maxShapes;
	if (maximum === Number.POSITIVE_INFINITY) {
		return counted;
	}
	const total = counted + countAuthoredElements(elements);
	if (total > maximum) {
		throw new PptxLoadLimitError(
			`PPTX presentation has at least ${count(total)} shapes, more than the maxShapes limit of ${count(maximum)}`,
			'maxShapes',
			maximum,
			total,
		);
	}
	return total;
}

/** The elements of one slide's tree that {@link countPptxShapesWithinLimits} counts. */
function countAuthoredElements(elements: readonly unknown[]): number {
	let total = 0;
	// One list and the next index in it per open group: no recursion, and the
	// stack only holds groups that were counted.
	const lists: (readonly unknown[])[] = [elements];
	const positions: number[] = [0];
	while (lists.length > 0) {
		const top = lists.length - 1;
		const list = lists[top]!;
		const position = positions[top]!;
		if (position >= list.length) {
			lists.pop();
			positions.pop();
			continue;
		}
		positions[top] = position + 1;
		const element = list[position];
		if (typeof element !== 'object' || element === null) {
			continue;
		}
		const { id, type, children } = element as { id?: unknown; type?: unknown; children?: unknown };
		if (typeof id === 'string' && isTemplateElementId(id)) {
			continue;
		}
		total++;
		if (type === 'group' && Array.isArray(children) && children.length > 0) {
			lists.push(children);
			positions.push(0);
		}
	}
	return total;
}
