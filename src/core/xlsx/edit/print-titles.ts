import {
	PRINT_TITLES_NAME,
	printTitles,
	printTitlesFormula,
	type PrintTitleSpan,
} from '../layout/print-titles';
import { type EditContext, sheetAt } from './context';
import { deleteDefinedName, setDefinedName } from './view';

export function setPrintTitles(
	ctx: EditContext,
	sheet: number,
	patch: { rows?: PrintTitleSpan | undefined; columns?: PrintTitleSpan | undefined },
): void {
	const ws = sheetAt(ctx.workbook, sheet);
	const existing = printTitles(ctx.workbook, sheet);
	if (existing.unsupported && !Object.hasOwn(patch, 'rows') && !Object.hasOwn(patch, 'columns'))
		return;
	if (existing.unsupported && (!Object.hasOwn(patch, 'rows') || !Object.hasOwn(patch, 'columns')))
		throw new Error('Replace both print title ranges to change unsupported titles.');
	const next = { ...existing, ...patch };
	const formula = printTitlesFormula(ws.name, next);
	if (formula) setDefinedName(ctx, { name: PRINT_TITLES_NAME, formula, localSheet: sheet });
	else deleteDefinedName(ctx, PRINT_TITLES_NAME, sheet);
}
