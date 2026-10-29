// Scenario Manager > Merge: copies the scenarios of another sheet of the workbook onto a sheet.
// Excel appends them in order; one whose name is taken is renamed "<name> <yyyy-mm-dd>" (measured
// with Scenarios.Merge on Excel 16), and the comment, protection and user come along.
import type { Scenario, ScenarioSet } from '../model';
import { type EditContext, sheetAt } from './context';
import { scenariosOf } from './scenarios';

/** The local date as `yyyy-mm-dd`. */
const isoDate = (date: Date): string =>
	`${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;

/** A name free among `taken`: `base`, else `base <date>`, else `base <date> 2`... */
function freeName(base: string, taken: Set<string>, date: string): string {
	if (!taken.has(base.toLowerCase())) return base;
	const dated = `${base} ${date}`;
	if (!taken.has(dated.toLowerCase())) return dated;
	for (let n = 2; ; n++) {
		const candidate = `${dated} ${n}`;
		if (!taken.has(candidate.toLowerCase())) return candidate;
	}
}

/**
 * Merges the scenarios of sheet `from` into sheet `s`; returns how many were added. One undo step.
 * The changing cells keep their addresses, so the sheets are expected to share a layout.
 */
export function mergeScenarios(
	ctx: EditContext,
	s: number,
	from: number,
	now: Date = new Date(),
): number {
	if (s === from) throw new RangeError('Choose a different sheet to merge scenarios from.');
	const sheet = sheetAt(ctx.workbook, s);
	const incoming = scenariosOf(sheetAt(ctx.workbook, from));
	if (!incoming.length) return 0;
	const taken = new Set(scenariosOf(sheet).map((scenario) => scenario.name.toLowerCase()));
	const date = isoDate(now);
	const added: Scenario[] = incoming.map((scenario) => {
		const name = freeName(scenario.name, taken, date);
		taken.add(name.toLowerCase());
		return { ...structuredClone(scenario), name };
	});
	ctx.run('Merge scenarios', 'view', [{ kind: 'parts', sheet: s, parts: ['scenarios'] }], () => {
		const set: ScenarioSet = (sheet.scenarios ??= { list: [] });
		set.list.push(...added);
	});
	return added.length;
}
