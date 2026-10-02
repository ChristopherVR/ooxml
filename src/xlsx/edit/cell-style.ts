// Applying a named cell style (Home > Cell Styles) to ranges.
import { type CellRange, normalizeRange } from '../address.js';
import { applyBuiltinStyle, type BuiltinCellStyle, builtinCellStyle } from '../cell-styles.js';
import type { CellStyle, Workbook } from '../model.js';
import { internStyle, styleAt } from '../styles.js';
import { type EditContext, sheetAt } from './context.js';
import { formatScope, patchRangeWith } from './format.js';

/** The named style entry, added to `workbook.namedStyles` from the catalogue when missing. */
function ensureNamedStyle(workbook: Workbook, name: string): Workbook['namedStyles'][number] {
	const lower = name.toLowerCase();
	const known = workbook.namedStyles.find((n) => n.name.toLowerCase() === lower);
	if (known) return known;
	const builtin = builtinCellStyle(name);
	if (!builtin) throw new Error(`Unknown cell style "${name}".`);
	const normal = styleAt(workbook, 0);
	const entry = {
		name: builtin.name,
		style: applyBuiltinStyle(normal, builtin, normal),
		builtinId: builtin.builtinId,
	};
	workbook.namedStyles.push(entry);
	return entry;
}

/** The style a cell gets when the named style is applied over `base`. */
function styled(
	base: CellStyle,
	named: Workbook['namedStyles'][number],
	builtin: BuiltinCellStyle | undefined,
	normal: CellStyle,
): CellStyle {
	if (!builtin) return { ...structuredClone(named.style), cellStyleName: named.name };
	// The workbook's own definition supplies the values of the aspects the style includes.
	const source = named.style;
	const aspects: BuiltinCellStyle = {
		...builtin,
		...(builtin.font ? { font: source.font } : {}),
		...(builtin.fill ? { fill: source.fill } : {}),
		...(builtin.border ? { border: source.border } : {}),
		...(builtin.numFmt !== undefined ? { numFmt: source.numFmt } : {}),
	};
	const next = applyBuiltinStyle(base, aspects, normal);
	next.cellStyleName = named.name;
	return next;
}

/**
 * Applies a named cell style (a built-in one such as `Good` or `Heading 1`, or one the workbook
 * defines) to ranges: the aspects the style includes replace the cells' formatting and the
 * cells record the style name, so the saved file lists it in Excel's gallery.
 */
export function applyCellStyle(
	ctx: EditContext,
	s: number,
	ranges: CellRange[],
	name: string,
): void {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	const normalized = ranges.map(normalizeRange);
	if (!normalized.length) return;
	const builtin = builtinCellStyle(name);
	if (!builtin && !workbook.namedStyles.some((n) => n.name.toLowerCase() === name.toLowerCase()))
		throw new Error(`Unknown cell style "${name}".`);
	ctx.run(
		`Cell style ${builtin?.name ?? name}`,
		'format',
		[formatScope(s, normalized)],
		() => {
			const named = ensureNamedStyle(workbook, name);
			const normal = styleAt(workbook, 0);
			const memo = new Map<number, number>();
			const patcher = (id: number | undefined): number => {
				const key = id ?? 0;
				let next = memo.get(key);
				if (next === undefined) {
					next = internStyle(workbook, styled(styleAt(workbook, id), named, builtin, normal));
					memo.set(key, next);
				}
				return next;
			};
			for (const range of normalized) patchRangeWith(sheet, range, patcher);
		},
		{ sheet: s, ranges: normalized },
	);
}
