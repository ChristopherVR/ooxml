// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The counters of numbered lists in document order, shared by list labels and LISTNUM fields.
import type { NumberingCatalog, NumberingLevelDefinition } from './numbering-model';
import { resolveNumberingLevel } from './numbering-parse';
import { formatListNumber } from './numbering-format';

/** Per `numId`, the current value of each level. */
export class NumberingCounters {
	private counts = new Map<string, Map<number, number>>();
	private started = new Map<string, Set<number>>();

	private forNum(numId: string): Map<number, number> {
		let counts = this.counts.get(numId);
		if (!counts) {
			counts = new Map();
			this.counts.set(numId, counts);
		}
		return counts;
	}

	private startedLevels(numId: string): Set<number> {
		let started = this.started.get(numId);
		if (!started) {
			started = new Set();
			this.started.set(numId, started);
		}
		return started;
	}

	/** Advances one level and resets deeper counters according to their resolved restart rules. */
	advance(
		numId: string,
		level: number,
		def: NumberingLevelDefinition,
		catalog: NumberingCatalog,
	): number {
		const counts = this.forNum(numId);
		const started = this.startedLevels(numId);
		const next = started.has(level) ? (counts.get(level) ?? def.start) + 1 : def.start;
		counts.set(level, next);
		started.add(level);
		for (const deeper of [...started]) {
			if (deeper <= level) continue;
			const deeperDef = resolveNumberingLevel(catalog, numId, deeper);
			const restart = deeperDef?.lvlRestart;
			if (restart === 0) continue;
			// OOXML uses a one-based trigger. Invalid deeper/self triggers are ignored,
			// leaving the default (previous level). Word also restarts on a skipped
			// higher level, independently of that intermediate level's own restart rule.
			const trigger =
				restart !== undefined && restart > 0 && restart <= deeper ? restart - 1 : deeper - 1;
			if (level > trigger) continue;
			counts.delete(deeper);
			started.delete(deeper);
		}
		// Nested items consume the initial value of omitted ancestors, even when
		// their marker has no ancestor placeholders. Later explicit items advance it.
		for (let ancestor = 0; ancestor < level; ancestor++) {
			if (started.has(ancestor)) continue;
			const ancestorDef = resolveNumberingLevel(catalog, numId, ancestor);
			if (!ancestorDef) continue;
			counts.set(ancestor, ancestorDef.start);
			started.add(ancestor);
		}
		return next;
	}

	/** Sets a level's value, as a LISTNUM `\s` start does, restarting the deeper levels. */
	set(numId: string, level: number, value: number): void {
		const counts = this.forNum(numId);
		const started = this.startedLevels(numId);
		counts.set(level, value);
		started.add(level);
		for (const deeper of [...started])
			if (deeper > level) {
				counts.delete(deeper);
				started.delete(deeper);
			}
	}

	/** Current value for an ancestor placeholder, defaulting to its own start when never instantiated. */
	current(numId: string, level: number, start: number): number {
		return this.forNum(numId).get(level) ?? start;
	}
}

/** A level's `w:lvlText` with each `%n` replaced by that level's current value. */
export function substituteLvlText(
	lvlText: string,
	numId: string,
	level: number,
	catalog: NumberingCatalog,
	counters: NumberingCounters,
	forceDecimal: boolean,
): string {
	return lvlText.replace(/%([1-9])/g, (match, digit: string) => {
		const ancestorLevel = Number(digit) - 1;
		if (ancestorLevel > level) return match;
		const ancestorDef = resolveNumberingLevel(catalog, numId, ancestorLevel);
		const start = ancestorDef?.start ?? 1;
		const value = counters.current(numId, ancestorLevel, start);
		const fmt = forceDecimal ? 'decimal' : (ancestorDef?.numFmt ?? 'decimal');
		return formatListNumber(fmt, value);
	});
}
