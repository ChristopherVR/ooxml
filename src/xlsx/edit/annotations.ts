import { type CellAddress, type CellRange, normalizeRange, rangesIntersect } from '../address.js';
import { getCell } from '../cells.js';
import type { ConditionalFormat, DataValidation, Hyperlink } from '../model.js';
import { writeValue } from './cell-values.js';
import { type EditContext, sheetAt } from './context.js';
import { patchRange } from './format.js';
import { subtractRange } from './range-math.js';

/** Adds, replaces or (with `text` undefined) removes the comment on a cell. */
export function setComment(
	ctx: EditContext,
	s: number,
	at: CellAddress,
	text: string | undefined,
	author: string,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const index = sheet.comments.findIndex(
		(c) => c.address.row === at.row && c.address.col === at.col,
	);
	if (index < 0 && text === undefined) return;
	ctx.run(
		text === undefined ? 'Delete comment' : index < 0 ? 'New comment' : 'Edit comment',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			if (text === undefined) {
				sheet.comments.splice(index, 1);
				return;
			}
			const existing = sheet.comments[index];
			if (existing) {
				existing.text = text;
				existing.author = author;
			} else sheet.comments.push({ address: { row: at.row, col: at.col }, author, text });
		},
		{ sheet: s, ranges: [{ start: at, end: at }] },
	);
}

/**
 * Sets (or removes) the hyperlink on a range. Links overlapping the range are replaced; an empty
 * anchor cell gets the link's display text and the cells get the hyperlink look.
 */
export function setHyperlink(
	ctx: EditContext,
	s: number,
	range: CellRange,
	link: Omit<Hyperlink, 'range'> | undefined,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	if (!link && !sheet.hyperlinks.some((h) => rangesIntersect(h.range, r))) return;
	if (link && !link.target && !link.location)
		throw new Error('A hyperlink needs a target or a location.');
	ctx.run(
		link ? 'Insert link' : 'Remove link',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.hyperlinks = sheet.hyperlinks.filter((h) => !rangesIntersect(h.range, r));
			if (!link) return;
			sheet.hyperlinks.push({ ...structuredClone(link), range: r });
			const anchor = getCell(sheet, r.start.row, r.start.col);
			if (!anchor || anchor.value === null || anchor.value === '')
				writeValue(
					sheet,
					r.start.row,
					r.start.col,
					link.display ?? link.target ?? link.location ?? '',
				);
			patchRange(ctx.workbook, sheet, r, { font: { color: { theme: 10 }, underline: 'single' } });
		},
		{ sheet: s, ranges: [r] },
	);
}

/** Adds a conditional format whose rules take the top priorities (as a new rule does in Excel). */
export function addConditionalFormat(ctx: EditContext, s: number, format: ConditionalFormat): void {
	const sheet = sheetAt(ctx.workbook, s);
	if (!format.ranges.length || !format.rules.length) return;
	ctx.run(
		'Conditional formatting',
		'format',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const count = format.rules.length;
			for (const cf of sheet.conditionalFormats)
				for (const rule of cf.rules) rule.priority += count;
			const added = structuredClone(format);
			added.ranges = added.ranges.map(normalizeRange);
			added.rules.forEach((rule, i) => {
				rule.priority = i + 1;
			});
			sheet.conditionalFormats.push(added);
		},
		{ sheet: s, ranges: format.ranges },
	);
}

/** Clears conditional formats from a range (rules covering more keep the rest) or the sheet. */
export function clearConditionalFormats(ctx: EditContext, s: number, range?: CellRange): void {
	const sheet = sheetAt(ctx.workbook, s);
	if (!sheet.conditionalFormats.length) return;
	const r = range && normalizeRange(range);
	ctx.run(
		'Clear rules',
		'format',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.conditionalFormats = r
				? sheet.conditionalFormats
						.map((cf) => ({ ...cf, ranges: cf.ranges.flatMap((cr) => subtractRange(cr, r)) }))
						.filter((cf) => cf.ranges.length > 0)
				: [];
		},
		{ sheet: s, ...(r ? { ranges: [r] } : {}) },
	);
}

/**
 * Sets the validation of a range (replacing what overlapped it) or clears it when undefined.
 * `showErrorMessage` and `showInputMessage` left out default to on, as Excel's Data Validation
 * dialog does (the file format's default for an absent attribute is off).
 */
export function setDataValidation(
	ctx: EditContext,
	s: number,
	validation: DataValidation | undefined,
	range: CellRange,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	ctx.run(
		validation ? 'Data validation' : 'Clear validation',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.dataValidations = sheet.dataValidations
				.map((dv) => ({ ...dv, ranges: dv.ranges.flatMap((dr) => subtractRange(dr, r)) }))
				.filter((dv) => dv.ranges.length > 0);
			if (validation)
				sheet.dataValidations.push({
					showErrorMessage: true,
					showInputMessage: true,
					...structuredClone(validation),
					ranges: [r],
				});
		},
		{ sheet: s, ranges: [r] },
	);
}
