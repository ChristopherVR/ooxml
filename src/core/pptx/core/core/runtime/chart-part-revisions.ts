// The states a classic chart part was saved in during a session. Chart data carries the state it
// was edited from (`PptxChartData.savedRevision`), so data that undo or redo brings back saves from
// the part it describes: undoing a removed row after a save gives that part back, range and cache
// together, instead of a difference applied to whatever was saved last.
//
// A state is held by the chart data edited from it, not by the handler: history entries copied with
// `cloneSlide` share it, and it goes when the last of them does. A copy that kept only its `id`
// (`structuredClone`, JSON) finds it while some data still holds it. It keeps only the part's XML,
// compressed; the chart data it describes is read back from that XML when it is needed. When the
// states still held pass `CHART_PART_REVISION_BUDGET` bytes, the oldest are released. Data edited
// from a state that is gone or released saves as data edited directly from the part as it is now
// (`source` returns `null`).
import { deflateSync, inflateSync, strFromU8, strToU8 } from 'fflate';
import type { PptxChartData, PptxChartSavedRevision } from '../../types';
import { chartPartSheets } from '../../utils/chart-workbook-sheets';

/** Compressed bytes of earlier saved states a handler keeps for undo before releasing the oldest. */
export const CHART_PART_REVISION_BUDGET = 32 * 1024 * 1024;

class ChartPartRevision implements PptxChartSavedRevision {
	readonly id: number;
	readonly #paths: Set<string>;
	#part: Uint8Array | undefined;
	/** The workbook sheets the kept part reads, lower case. */
	#sheets: readonly string[] = [];

	constructor(id: number, path: string) {
		this.id = id;
		this.#paths = new Set([path]);
		// A frozen object is never wrapped in a framework's reactive proxy, which could not reach
		// the private fields.
		Object.freeze(this);
	}

	get size(): number {
		return this.#part?.byteLength ?? 0;
	}

	belongsTo(path: string): boolean {
		return this.#paths.has(path);
	}

	share(path: string): void {
		this.#paths.add(path);
	}

	/** Keeps `xml`, the part as this state saved it, once a later save replaces it. */
	keep(xml: string): boolean {
		if (this.#part) return false;
		this.#part = deflateSync(strToU8(xml), { level: 1 });
		this.#sheets = [...chartPartSheets(xml)];
		return true;
	}

	get sheets(): readonly string[] {
		return this.#sheets;
	}

	xml(): string | undefined {
		return this.#part && strFromU8(inflateSync(this.#part));
	}

	release(): void {
		this.#part = undefined;
		this.#sheets = [];
	}
}

export class ChartPartRevisions {
	readonly #latest = new Map<string, ChartPartRevision>();
	readonly #saved = new Set<string>();
	/** Every state by id, in the order they were made. */
	readonly #states = new Map<number, WeakRef<ChartPartRevision>>();
	readonly #budget: number;
	// Never reused, so data from an earlier load never names a state of this one.
	#ids = 0;

	constructor(budget = CHART_PART_REVISION_BUDGET) {
		this.#budget = budget;
	}

	/** Marks `data`, as loaded, as describing the part at `path` as it is now. */
	start(path: string, data: PptxChartData): void {
		const revision = new ChartPartRevision(++this.#ids, path);
		this.#states.set(revision.id, new WeakRef(revision));
		this.#latest.set(path, revision);
		data.savedRevision = revision;
	}

	/**
	 * The part XML `data` was edited from: `undefined` when it is the part as it is now (or nothing
	 * was saved to `path` yet), `null` when that state is no longer kept.
	 */
	source(path: string, data: PptxChartData): string | null | undefined {
		const latest = this.#latest.get(path);
		const from = data.savedRevision;
		if (!latest || !this.#saved.has(path) || from?.id === latest.id) return undefined;
		const state =
			from instanceof ChartPartRevision ? from : from && this.#states.get(from.id)?.deref();
		return state?.belongsTo(path) ? (state.xml() ?? null) : null;
	}

	/**
	 * Records that `data` now describes the part at `path`; `previous` is the part as it was, kept
	 * for the data still edited from that state.
	 */
	record(path: string, previous: string, data: PptxChartData): void {
		const replaced = this.#latest.get(path);
		if (replaced?.keep(previous)) this.#trim();
		this.start(path, data);
		this.#saved.add(path);
		delete data.structureEdits;
	}

	/** Gives `destination`, a copy of the part at `source`, the state data of both now shares. */
	copy(source: string, destination: string): void {
		const latest = this.#latest.get(source);
		if (!latest) return;
		latest.share(destination);
		this.#latest.set(destination, latest);
		if (this.#saved.has(source)) this.#saved.add(destination);
	}

	delete(path: string): void {
		this.#latest.delete(path);
		this.#saved.delete(path);
	}

	clear(): void {
		this.#latest.clear();
		this.#saved.clear();
		this.#states.clear();
	}

	/** The workbook sheets the kept states of the part at `path` read, lower case. */
	sheetsHeld(path: string): Set<string> {
		return new Set(
			this.#held().flatMap((revision) => (revision.belongsTo(path) ? revision.sheets : [])),
		);
	}

	/** Compressed bytes the kept states hold. */
	get size(): number {
		return this.#held().reduce((total, revision) => total + revision.size, 0);
	}

	/** The states still held that keep a part, oldest first; forgets the ones that are gone. */
	#held(): ChartPartRevision[] {
		const held: ChartPartRevision[] = [];
		for (const [id, reference] of this.#states) {
			const revision = reference.deref();
			if (!revision) this.#states.delete(id);
			else if (revision.size) held.push(revision);
		}
		return held;
	}

	/** Releases the oldest kept states until the rest fit the budget. */
	#trim(): void {
		const held = this.#held();
		let total = held.reduce((sum, revision) => sum + revision.size, 0);
		for (const revision of held) {
			if (total <= this.#budget) return;
			total -= revision.size;
			revision.release();
		}
	}
}
