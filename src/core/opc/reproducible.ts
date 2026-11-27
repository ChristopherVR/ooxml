/**
 * What a save takes from the machine, made injectable so a caller can reproduce its bytes.
 *
 * Office stamps the time of a save (`dcterms:modified`, Visio's `TimeRefreshed`) and draws some
 * ids at random (table `a16:colId`, SmartArt model ids, data store item ids, content control
 * ids). Every writer keeps that behaviour by default, and every writer's save options take the
 * same {@link ReproducibleSaveOptions}: `now` fixes the clock and `ids` the id source, so two
 * saves of the same input with the same `now` and a fresh `seededIdSource(seed)` each are
 * byte-identical. ZIP entries are dated 1980-01-01 00:00:00, as Office writes them, unless
 * `zipDate` asks otherwise.
 *
 * @module opc/reproducible
 */
import type JSZip from 'jszip';

/** A source of unsigned 32-bit integers that a save draws its random ids from. */
export interface IdSource {
	/** The next unsigned 32-bit integer (0 to 2^32 - 1). */
	next(): number;
}

/** The options every writer's save takes to make its output reproducible. */
export interface ReproducibleSaveOptions {
	/**
	 * The time the save stamps where Office stamps one (the core properties' modified date, a
	 * new package's created date, a Visio data recordset's refresh time). Default: the time of
	 * the save.
	 */
	now?: Date;
	/** Where the ids Office draws at random come from. Default: {@link randomIdSource}. */
	ids?: IdSource;
	/** The date of every ZIP entry. Default: {@link OFFICE_ZIP_DATE}, as Office writes. */
	zipDate?: Date;
}

/** 1980-01-01 00:00:00 UTC, the DOS epoch: the date Office gives every entry of a package it saves. */
export const OFFICE_ZIP_DATE: Date = new Date(Date.UTC(1980, 0, 1));

/** Ids from the platform's cryptographic random source (falling back to `Math.random`). */
export const randomIdSource: IdSource = {
	next(): number {
		const crypto = globalThis.crypto;
		if (crypto?.getRandomValues) return crypto.getRandomValues(new Uint32Array(1))[0] ?? 0;
		return Math.floor(Math.random() * 0x1_0000_0000);
	},
};

/**
 * A reproducible id source: the same seed always gives the same sequence (mulberry32, identical
 * on every engine). Make a fresh one per save, since a source moves on with every id it gives.
 */
export function seededIdSource(seed: number): IdSource {
	let state = seed >>> 0;
	return {
		next(): number {
			state = (state + 0x6d2b79f5) >>> 0;
			let t = state;
			t = Math.imul(t ^ (t >>> 15), t | 1);
			t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
			return (t ^ (t >>> 14)) >>> 0;
		},
	};
}

/** An integer from `min` to `max` (both included, at most 2^32 values) drawn from `ids`. */
export function idInRange(ids: IdSource, min: number, max: number): number {
	return min + (ids.next() % (max - min + 1));
}

const hex8 = (value: number) => (value >>> 0).toString(16).toUpperCase().padStart(8, '0');

/** A braced, upper-case version 4 GUID (`{XXXXXXXX-XXXX-4XXX-YXXX-XXXXXXXXXXXX}`) from `ids`. */
export function guidFrom(ids: IdSource): string {
	const a = hex8(ids.next());
	const b = hex8(ids.next());
	const c = hex8(ids.next());
	const d = hex8(ids.next());
	const variant = (8 + (Number.parseInt(c[0] ?? '0', 16) % 4)).toString(16).toUpperCase();
	return `{${a}-${b.slice(0, 4)}-4${b.slice(5)}-${variant}${c.slice(1, 4)}-${c.slice(4)}${d}}`;
}

/** The clock, id source and ZIP date of one save, with the defaults filled in. */
export interface SaveEnvironment {
	now: Date;
	ids: IdSource;
	zipDate: Date;
	/**
	 * The body (minor) font of the document's theme, for the parts a save creates that name it
	 * (a new chart's embedded workbook). Unset: `OFFICE_THEME_FONTS.minor` (drawingml).
	 */
	minorFont?: string;
}

function validDate(date: Date | undefined, name: string): Date | undefined {
	if (date === undefined) return undefined;
	if (!(date instanceof Date) || Number.isNaN(date.getTime()))
		throw new RangeError(`The save option ${name} must be a valid Date`);
	return date;
}

/**
 * The environment of a save from its options: `now` defaults to the current time (read once,
 * so every stamp of one save agrees), `ids` to {@link randomIdSource}, `zipDate` to
 * {@link OFFICE_ZIP_DATE}. An invalid date throws rather than saving a broken one.
 */
export function saveEnvironment(options?: ReproducibleSaveOptions): SaveEnvironment {
	return {
		now: validDate(options?.now, 'now') ?? new Date(),
		ids: options?.ids ?? randomIdSource,
		zipDate: validDate(options?.zipDate, 'zipDate') ?? OFFICE_ZIP_DATE,
	};
}

const bound = new WeakMap<object, SaveEnvironment>();

/**
 * Binds `environment` to the package a save is writing (or an XML document of it), so the writers deep inside that save
 * (chart workbooks, SmartArt, content controls, data store items) read the same clock and id
 * source without each call passing them on. The binding lives as long as the package object,
 * so concurrent saves of different packages never share one.
 */
export function bindSaveEnvironment<T extends object>(target: T, environment: SaveEnvironment): T {
	bound.set(target, environment);
	return target;
}

/** The environment bound to `target` (a package or one of its XML documents) by its save, else the defaults (now, random ids, 1980). */
export function saveEnvironmentOf(target: object | undefined): SaveEnvironment {
	const environment = target && bound.get(target);
	if (environment) return environment;
	const fresh = saveEnvironment();
	if (target) bound.set(target, fresh);
	return fresh;
}

/** Dates every entry of `zip` (call it just before generating the package). */
export function dateZipEntries(zip: JSZip, date: Date = OFFICE_ZIP_DATE): void {
	for (const entry of Object.values(zip.files)) entry.date = date;
}
