// Scenario Manager (Data > What-If Analysis): named sets of values for changing cells, stored as
// the sheet's `<scenarios>` element. Showing a scenario writes its values into the cells.
import {
	type CellAddress,
	formatAddress,
	normalizeRange,
	parseAddress,
	parseRange,
} from '../address';
import { getCell } from '../cells';
import type { Scenario, ScenarioCell, ScenarioSet, Workbook, Worksheet } from '../model';
import { type EditContext, sheetAt } from './context';
import { writeInput } from './cell-values';
import { formulaBarText } from './input-text';

/** Excel allows at most 32 changing cells in a scenario. */
export const MAX_SCENARIO_CELLS = 32;

export type ScenarioProblem =
	| 'name-empty'
	| 'name-exists'
	| 'name-too-long'
	| 'no-cells'
	| 'bad-cells'
	| 'too-many-cells'
	| 'locked'
	| 'missing';

const MESSAGES: Record<ScenarioProblem, string> = {
	'name-empty': 'Enter a name for the scenario.',
	'name-exists': 'A scenario with that name already exists.',
	'name-too-long': 'A scenario name can have at most 255 characters.',
	'no-cells': 'Enter the changing cells.',
	'bad-cells': 'The changing cells reference is not valid.',
	'too-many-cells': 'A scenario can have at most 32 changing cells.',
	locked: 'The scenario is protected and cannot be changed.',
	missing: 'The scenario does not exist.',
};

export const scenarioMessage = (problem: ScenarioProblem): string => MESSAGES[problem];

class ScenarioError extends Error {
	constructor(readonly problem: ScenarioProblem) {
		super(MESSAGES[problem]);
	}
}

/** Expands `B1,B3:B4` (an optional sheet prefix and `$` are ignored) into cell addresses. */
export function expandCells(text: string): CellAddress[] | undefined {
	const out: CellAddress[] = [];
	for (const part of text.split(/[,;]/)) {
		const ref = part.trim().replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '');
		if (!ref) continue;
		const range = parseRange(ref);
		if (!range) return undefined;
		const r = normalizeRange(range);
		for (let row = r.start.row; row <= r.end.row; row++)
			for (let col = r.start.col; col <= r.end.col; col++) {
				out.push({ row, col });
				if (out.length > 4096) return out;
			}
	}
	return out;
}

/** The scenarios of a sheet (empty when it has none). */
export const scenariosOf = (sheet: Worksheet): Scenario[] => sheet.scenarios?.list ?? [];

/** The changing-cell addresses of a scenario, in order. */
export function scenarioAddresses(scenario: Scenario): CellAddress[] {
	return scenario.cells.flatMap((c) => {
		const address = parseAddress(c.ref);
		return address ? [address] : [];
	});
}

function checkName(sheet: Worksheet, name: string, except?: number): string {
	const trimmed = name.trim();
	if (!trimmed) throw new ScenarioError('name-empty');
	if (trimmed.length > 255) throw new ScenarioError('name-too-long');
	const taken = scenariosOf(sheet).some(
		(s, i) => i !== except && s.name.toLowerCase() === trimmed.toLowerCase(),
	);
	if (taken) throw new ScenarioError('name-exists');
	return trimmed;
}

function cellsOf(addresses: CellAddress[]): CellAddress[] {
	if (!addresses.length) throw new ScenarioError('no-cells');
	if (addresses.length > MAX_SCENARIO_CELLS) throw new ScenarioError('too-many-cells');
	return addresses;
}

const textAt = (workbook: Workbook, sheet: Worksheet, a: CellAddress): string =>
	formulaBarText(workbook, getCell(sheet, a.row, a.col), { quote: false });

export interface ScenarioDraft {
	name: string;
	/** The changing cells, for example `B1,B3:B4`. */
	cells: string;
	comment?: string;
	locked?: boolean;
	hidden?: boolean;
	user?: string;
	/** Values by A1 reference; missing entries take the cell's current value. */
	values?: Record<string, string>;
}

function build(ctx: EditContext, sheet: Worksheet, draft: ScenarioDraft, name: string): Scenario {
	const expanded = expandCells(draft.cells);
	if (!expanded) throw new ScenarioError('bad-cells');
	const cells: ScenarioCell[] = cellsOf(expanded).map((a) => ({
		ref: formatAddress(a),
		value: draft.values?.[formatAddress(a)] ?? textAt(ctx.workbook, sheet, a),
	}));
	const scenario: Scenario = { name, cells };
	if (draft.comment) scenario.comment = draft.comment;
	if (draft.locked) scenario.locked = true;
	if (draft.hidden) scenario.hidden = true;
	if (draft.user) scenario.user = draft.user;
	return scenario;
}

/**
 * Validates a draft's name and changing cells (throwing the same errors as adding it) and
 * returns the values the Scenario Values dialog starts with: those stored in scenario `except`
 * (when editing) else the cells' current contents.
 */
export function previewScenario(
	workbook: Workbook,
	s: number,
	draft: ScenarioDraft,
	except?: number,
): { ref: string; value: string }[] {
	const sheet = sheetAt(workbook, s);
	checkName(sheet, draft.name, except);
	const expanded = expandCells(draft.cells);
	if (!expanded) throw new ScenarioError('bad-cells');
	const stored = except === undefined ? undefined : scenariosOf(sheet)[except];
	return cellsOf(expanded).map((a) => {
		const ref = formatAddress(a);
		return {
			ref,
			value: stored?.cells.find((c) => c.ref === ref)?.value ?? textAt(workbook, sheet, a),
		};
	});
}

const PARTS = (s: number) => [{ kind: 'parts' as const, sheet: s, parts: ['scenarios' as const] }];

/** Adds a scenario that records the changing cells' values (or `draft.values`). */
export function addScenario(ctx: EditContext, s: number, draft: ScenarioDraft): number {
	const sheet = sheetAt(ctx.workbook, s);
	const name = checkName(sheet, draft.name);
	const scenario = build(ctx, sheet, draft, name);
	return ctx.run('Add scenario', 'view', PARTS(s), () => {
		const set: ScenarioSet = (sheet.scenarios ??= { list: [] });
		set.list.push(scenario);
		return set.list.length - 1;
	});
}

/** Replaces a scenario's name, comment and cell values. */
export function editScenario(
	ctx: EditContext,
	s: number,
	index: number,
	draft: ScenarioDraft,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const existing = Number.isInteger(index) ? scenariosOf(sheet)[index] : undefined;
	if (!existing) throw new ScenarioError('missing');
	if (existing.locked && sheet.protection) throw new ScenarioError('locked');
	const name = checkName(sheet, draft.name, index);
	const scenario = build(ctx, sheet, draft, name);
	ctx.run('Edit scenario', 'view', PARTS(s), () => {
		(sheet.scenarios as ScenarioSet).list[index] = scenario;
	});
}

export function deleteScenario(ctx: EditContext, s: number, index: number): void {
	const sheet = sheetAt(ctx.workbook, s);
	const existing = Number.isInteger(index) ? scenariosOf(sheet)[index] : undefined;
	if (!existing) throw new ScenarioError('missing');
	if (existing.locked && sheet.protection) throw new ScenarioError('locked');
	ctx.run('Delete scenario', 'view', PARTS(s), () => {
		const set = sheet.scenarios as ScenarioSet;
		set.list.splice(index, 1);
		const fix = (n: number | undefined): number | undefined =>
			n === undefined ? undefined : n === index ? undefined : n > index ? n - 1 : n;
		const current = fix(set.current);
		const show = fix(set.show);
		if (current === undefined) delete set.current;
		else set.current = current;
		if (show === undefined) delete set.show;
		else set.show = show;
		if (!set.list.length) delete sheet.scenarios;
	});
}

/** Shows a scenario: writes its values into the changing cells. One undo step. */
export function showScenario(ctx: EditContext, s: number, index: number): void {
	const sheet = sheetAt(ctx.workbook, s);
	const scenario = scenariosOf(sheet)[index];
	if (!scenario) throw new ScenarioError('missing');
	const addresses = scenarioAddresses(scenario);
	const ranges = addresses.map((a) => ({ start: a, end: a }));
	ctx.run(
		'Show scenario',
		'cells',
		[...ranges.map((r) => ({ kind: 'cells' as const, sheet: s, ranges: [r] })), ...PARTS(s)],
		() => {
			for (const cell of scenario.cells) {
				const a = parseAddress(cell.ref);
				if (a && !cell.deleted) writeInput(ctx, sheet, a.row, a.col, cell.value);
			}
			const set = sheet.scenarios as ScenarioSet;
			set.show = index;
			set.current = index;
		},
		{ sheet: s, ranges },
	);
}
