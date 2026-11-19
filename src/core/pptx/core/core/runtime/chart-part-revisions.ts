// The states a classic chart part was saved in during a session. Chart data carries the revision it
// was edited from (`PptxChartData.savedRevision`), so data that undo or redo brings back saves from
// the part it describes: undoing a removed row after a save gives that part back, range and cache
// together, instead of a difference applied to whatever was saved last.
import type { PptxChartData } from '../../types';

/** A saved state of a chart part: its XML and the chart data that describes it, as JSON. */
export interface ChartPartRevision {
	xml: string;
	data?: string | undefined;
}

/** Saved states kept per part; data edited from an older one saves from the part as it is now. */
const KEPT_REVISIONS = 100;

interface PartHistory {
	latest: number;
	saved: Map<number, ChartPartRevision>;
}

export class ChartPartRevisions {
	readonly #parts = new Map<string, PartHistory>();

	/**
	 * The saved state `data` was edited from: `undefined` when it is the part as it is now, `null`
	 * when that state is no longer kept.
	 */
	source(path: string, data: PptxChartData): ChartPartRevision | null | undefined {
		const history = this.#parts.get(path);
		const base = data.savedRevision ?? 0;
		if (!history || base === history.latest) return undefined;
		return history.saved.get(base) ?? null;
	}

	/**
	 * Records `xml` as the newest state of `path`, saved for `data`, which is marked as describing
	 * it. `current` is the state the part had until now, kept the first time the part is saved.
	 */
	record(path: string, current: ChartPartRevision, xml: string, data: PptxChartData): string {
		let history = this.#parts.get(path);
		if (!history) {
			history = { latest: 0, saved: new Map([[0, current]]) };
			this.#parts.set(path, history);
		}
		history.latest += 1;
		data.savedRevision = history.latest;
		delete data.structureEdits;
		const json = JSON.stringify(data);
		history.saved.set(history.latest, { xml, data: json });
		for (const revision of history.saved.keys())
			if (history.saved.size > KEPT_REVISIONS) history.saved.delete(revision);
		return json;
	}

	/** Gives `destination`, a copy of `source`, the same saved states. */
	copy(source: string, destination: string, remap: (data: string) => string): void {
		const history = this.#parts.get(source);
		if (!history) return;
		const saved = new Map(
			[...history.saved].map(([revision, state]) => [
				revision,
				{ xml: state.xml, data: state.data === undefined ? undefined : remap(state.data) },
			]),
		);
		this.#parts.set(destination, { latest: history.latest, saved });
	}

	delete(path: string): void {
		this.#parts.delete(path);
	}

	clear(): void {
		this.#parts.clear();
	}
}
