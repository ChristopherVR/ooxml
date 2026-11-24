// Data > Subtotal: insert group SUBTOTAL formulas and outline levels, or remove existing subtotals.
import {
	columnLabel,
	getCell,
	type RemoveSubtotalsResult,
	type SubtotalFunction,
	type SubtotalResult,
} from 'ooxml-core/xlsx';
import { guessHeader, regionOf, target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, el, field, fieldset, select } from './fields';
import { button, showDialog } from './frame';

/** Shared aggregation function choices for Subtotal and Consolidate dialogs. */
export const SUBTOTAL_FUNCTION_OPTIONS: ReadonlyArray<readonly [SubtotalFunction, string]> = [
	['sum', 'Sum'],
	['counta', 'Count'],
	['average', 'Average'],
	['max', 'Max'],
	['min', 'Min'],
	['product', 'Product'],
	['count', 'Count Numbers'],
	['stdev', 'StdDev'],
	['stdevp', 'StdDevp'],
	['var', 'Var'],
	['varp', 'Varp'],
];

export async function openSubtotal(
	ctx: EditorContext,
): Promise<SubtotalResult | RemoveSubtotalsResult | undefined> {
	const t = target(ctx);
	if (!t) return undefined;
	const region = regionOf(t);
	const hasHeader = guessHeader(t.ws, region);
	const columns: { col: number; label: string }[] = [];
	for (let col = region.start.col; col <= region.end.col; col++) {
		const head = getCell(t.ws, region.start.row, col)?.value;
		const label =
			hasHeader && head !== null && head !== undefined && head !== ''
				? String(typeof head === 'object' ? head.error : head)
				: ctx.t('Column {name}', { name: columnLabel(col) });
		columns.push({ col, label });
	}

	const groupBy = select(
		ctx,
		columns.map((c) => [String(c.col), c.label] as const),
		String(region.start.col),
		true,
	);
	const fn = select(ctx, SUBTOTAL_FUNCTION_OPTIONS, 'sum');
	const list = el(ctx, 'div');
	const checks: { col: number; input: HTMLInputElement }[] = [];
	for (const { col, label } of columns) {
		const box = el(ctx, 'label', 'xve-check');
		const input = el(ctx, 'input');
		input.type = 'checkbox';
		input.checked = col === region.end.col;
		input.setAttribute('aria-label', label);
		const span = el(ctx, 'span');
		span.textContent = label;
		box.append(input, span);
		list.append(box);
		checks.push({ col, input });
	}
	const replace = checkbox(ctx, 'Replace current subtotals', true);
	const pageBreak = checkbox(ctx, 'Page break between groups', false);
	const summaryBelow = checkbox(ctx, 'Summary below data', true);
	const removeAll = button(ctx, 'Remove All');
	let removed: RemoveSubtotalsResult | undefined;

	const result = await showDialog<SubtotalResult>(
		ctx,
		{
			name: 'subtotal',
			heading: 'Subtotal',
			body: [
				field(ctx, 'At each change in:', groupBy),
				field(ctx, 'Use function:', fn),
				fieldset(ctx, 'Add subtotal to:', list),
				replace.wrapper,
				pageBreak.wrapper,
				summaryBelow.wrapper,
			],
			buttons: [removeAll],
			opened: () => groupBy.focus(),
			submit: () => {
				const chosen = checks.filter((c) => c.input.checked).map((c) => c.col);
				if (!chosen.length) {
					ctx.toast(ctx.t('Select at least one column.'), 'warning');
					return undefined;
				}
				return t.session.applySubtotal(t.sheet, region, {
					atEachChangeIn: Number(groupBy.value),
					fn: fn.value as SubtotalFunction,
					addSubtotalTo: chosen,
					replaceExisting: replace.input.checked,
					pageBreakBetweenGroups: pageBreak.input.checked,
					summaryBelowData: summaryBelow.input.checked,
					hasHeader,
				});
			},
		},
		(open) => {
			removeAll.addEventListener('click', () => {
				removed = t.session.removeSubtotals(t.sheet, region, hasHeader);
				open.close();
			});
		},
	);
	return result ?? removed;
}
