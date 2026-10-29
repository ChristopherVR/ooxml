// Scenario Manager > Summary and Merge: the report on a new sheet and copying the scenarios of
// another sheet of the workbook.
import { scenariosOf } from 'ooxml-core/xlsx';
import { createScenarioSummary, target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { field, radios, select, text, textInput } from './fields';
import { showDialog } from './frame';

type Target = NonNullable<ReturnType<typeof target>>;

const report = (ctx: EditorContext, error: unknown): void =>
	ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'warning');

/** Creates the Scenario Summary sheet; returns whether one was created. */
export async function openScenarioSummary(ctx: EditorContext, t: Target): Promise<boolean> {
	const kind = radios(ctx, 'Report type', [['summary', 'Scenario summary']], 'summary');
	const results = textInput(ctx, t.ws.scenarios?.sqref ?? '');
	const created = await showDialog<number>(ctx, {
		name: 'scenario-summary',
		heading: 'Scenario Summary',
		body: [
			kind.element,
			field(ctx, 'Result cells:', results),
			text(ctx, 'Result cells are optional. Separate several references with commas.'),
		],
		opened: () => results.focus(),
		submit: () => {
			try {
				return createScenarioSummary(t.session, { sheet: t.sheet, resultCells: results.value });
			} catch (error) {
				report(ctx, error);
				return undefined;
			}
		},
	});
	return created !== undefined;
}

/** Copies the scenarios of another sheet onto this one; returns how many were added. */
export async function openScenarioMerge(ctx: EditorContext, t: Target): Promise<number> {
	const sources = t.workbook.sheets.flatMap((sheet, index) =>
		index !== t.sheet && scenariosOf(sheet).length ? [[String(index), sheet.name] as const] : [],
	);
	if (!sources.length) {
		ctx.toast(ctx.t('No other sheet has scenarios to merge.'), 'warning');
		return 0;
	}
	const sheet = select(ctx, sources, sources[0]![0], true);
	const added = await showDialog<number>(ctx, {
		name: 'scenario-merge',
		heading: 'Merge Scenarios',
		body: [field(ctx, 'Merge scenarios from sheet:', sheet)],
		opened: () => sheet.focus(),
		submit: () => {
			try {
				return t.session.mergeScenarios(t.sheet, Number(sheet.value));
			} catch (error) {
				report(ctx, error);
				return undefined;
			}
		},
	});
	return added ?? 0;
}
