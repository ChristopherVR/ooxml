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
import { flattenElementsDeep } from '../utils/flatten-elements';
import { DEFAULT_MAX_UNCOMPRESSED_BYTES, MAX_ZIP_ENTRY_COUNT, ZipBombError } from './types';
import type { PptxHandlerLoadOptions } from './types';

/** Limits on one presentation load; each field left out keeps its default. */
export interface PptxLoadLimits {
	/** Largest accepted input, in bytes, as passed to `load()` (default: no limit). */
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
	 * inside a group). Layout and master elements are not counted (default: no limit).
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
	 * The observed value: exact for `maxInputBytes` and `maxSlides`; for
	 * `maxShapes`, the count when the load stopped (the deck has at least this many).
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
 * Adds one slide's parsed elements (group children included) to `counted` and
 * fails as soon as the total passes `maxShapes`. Returns the new total.
 */
export function countPptxShapesWithinLimits(
	counted: number,
	elements: readonly PptxElement[],
	limits: Required<PptxLoadLimits>,
): number {
	if (limits.maxShapes === Number.POSITIVE_INFINITY) {
		return counted;
	}
	const total = counted + flattenElementsDeep(elements).length;
	if (total > limits.maxShapes) {
		throw new PptxLoadLimitError(
			`PPTX presentation has at least ${count(total)} shapes, more than the maxShapes limit of ${count(limits.maxShapes)}`,
			'maxShapes',
			limits.maxShapes,
			total,
		);
	}
	return total;
}
