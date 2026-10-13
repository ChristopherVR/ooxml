import { MAX_COL, MAX_ROW, columnLabel, quoteSheetName } from '../address';
import { parseFormula } from '../formula/parser';
import type { FormulaAst } from '../formula/ast';
import type { Workbook } from '../model';

export interface PrintTitleSpan {
	start: number;
	end: number;
}
export interface PrintTitles {
	rows?: PrintTitleSpan;
	columns?: PrintTitleSpan;
	unsupported?: boolean;
}
export const PRINT_TITLES_NAME = '_xlnm.Print_Titles';

function refs(ast: FormulaAst): FormulaAst[] {
	return ast.type === 'binary' && ast.op === ',' ? [...refs(ast.left), ...refs(ast.right)] : [ast];
}

/** Reads the native scoped name without replacing unsupported references. */
export function printTitles(workbook: Workbook, sheetIndex: number): PrintTitles {
	const name = workbook.definedNames.find(
		(entry) => entry.name === PRINT_TITLES_NAME && entry.localSheet === sheetIndex,
	);
	if (!name) return {};
	try {
		const result: PrintTitles = {};
		for (const node of refs(parseFormula(`(${name.formula})`))) {
			if (
				node.type !== 'ref' ||
				!node.ref ||
				node.prefix?.book ||
				node.prefix?.sheet2 ||
				node.spill ||
				(node.prefix &&
					node.prefix.sheet.toLowerCase() !== workbook.sheets[sheetIndex]?.name.toLowerCase())
			)
				throw new Error('Unsupported print title reference');
			const key =
				node.ref.kind === 'rows' ? 'rows' : node.ref.kind === 'cols' ? 'columns' : undefined;
			if (!key || result[key]) throw new Error('Unsupported print title reference');
			const axis = key === 'rows' ? 'row' : 'col';
			result[key] = {
				start: Math.min(node.ref.start[axis], node.ref.end[axis]),
				end: Math.max(node.ref.start[axis], node.ref.end[axis]),
			};
		}
		return result;
	} catch {
		return { unsupported: true };
	}
}

export function parsePrintTitleSpan(
	value: string,
	axis: 'row' | 'col',
): PrintTitleSpan | undefined {
	if (!value.trim()) return undefined;
	const ast = parseFormula(value.trim().replace(/^=/, ''));
	if (
		ast.type !== 'ref' ||
		!ast.ref ||
		ast.prefix ||
		ast.spill ||
		ast.ref.kind !== (axis === 'row' ? 'rows' : 'cols')
	)
		throw new Error('Print titles must be whole rows or whole columns.');
	return {
		start: Math.min(ast.ref.start[axis], ast.ref.end[axis]),
		end: Math.max(ast.ref.start[axis], ast.ref.end[axis]),
	};
}

export function formatPrintTitleSpan(span: PrintTitleSpan, axis: 'row' | 'col'): string {
	return axis === 'row'
		? `$${span.start + 1}:$${span.end + 1}`
		: `$${columnLabel(span.start)}:$${columnLabel(span.end)}`;
}

export function printTitlesFormula(
	sheet: string,
	titles: { rows?: PrintTitleSpan | undefined; columns?: PrintTitleSpan | undefined },
): string {
	return (['rows', 'columns'] as const)
		.flatMap((key) => {
			const span = titles[key];
			if (!span) return [];
			const max = key === 'rows' ? MAX_ROW : MAX_COL;
			if (
				![span.start, span.end].every(Number.isInteger) ||
				span.start < 0 ||
				span.end < span.start ||
				span.end > max
			)
				throw new Error('Print titles must be whole rows or whole columns.');
			return [
				`${quoteSheetName(sheet)}!${formatPrintTitleSpan(span, key === 'rows' ? 'row' : 'col')}`,
			];
		})
		.join(',');
}
