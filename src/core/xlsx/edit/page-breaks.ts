import { MAX_COL, MAX_ROW } from '../address';
import { NS, buildXml, children, elements, parseXml } from '../../xml/index';
import type { PrintAxis } from '../layout/print-breaks';
import { type EditContext, sheetAt } from './context';

/** Edits existing XML in place, preserving unknown attributes and non-manual breaks. */
function patch(
	xml: string,
	before: number | undefined,
	present: boolean,
	axis: PrintAxis,
): string | undefined {
	const doc = parseXml(xml);
	const root = doc.documentElement;
	const breaks = children(root, 'brk', NS.x);
	for (const node of breaks)
		if (
			!present &&
			/^(1|true)$/.test(node.getAttribute('man') ?? '') &&
			(before === undefined || Number(node.getAttribute('id')) === before)
		)
			root.removeChild(node);
	if (present && before !== undefined) {
		const existing = children(root, 'brk', NS.x).find(
			(node) => Number(node.getAttribute('id')) === before,
		);
		const node = existing ?? doc.createElementNS(NS.x, 'brk');
		node.setAttribute('id', String(before));
		node.setAttribute('min', '0');
		node.setAttribute('max', String(axis === 'row' ? MAX_COL : MAX_ROW));
		node.setAttribute('man', '1');
		if (!existing) root.appendChild(node);
	}
	const remaining = children(root, 'brk', NS.x);
	if (remaining.filter((node) => /^(1|true)$/.test(node.getAttribute('man') ?? '')).length > 1026)
		throw new Error('A worksheet can have at most 1026 manual breaks on each axis.');
	if (!elements(root).length) return undefined;
	root.setAttribute('count', String(remaining.length));
	root.setAttribute(
		'manualBreakCount',
		String(remaining.filter((node) => /^(1|true)$/.test(node.getAttribute('man') ?? '')).length),
	);
	return buildXml(root);
}

export function setPageBreak(
	ctx: EditContext,
	s: number,
	axis: PrintAxis,
	before: number,
	present = true,
): void {
	if (
		typeof present !== 'boolean' ||
		!['row', 'col'].includes(axis) ||
		!Number.isInteger(before) ||
		before <= 0 ||
		before > (axis === 'row' ? MAX_ROW : MAX_COL)
	)
		throw new RangeError('Choose a row or column after the first to set a page break.');
	const sheet = sheetAt(ctx.workbook, s);
	const key = axis === 'row' ? 'rowBreaks' : 'colBreaks';
	const source = sheet.preserved.get(key) ?? [`<${key} xmlns="${NS.x}"/>`];
	const result = source.flatMap((xml, index) => {
		const next = patch(xml, before, present && index === 0, axis);
		return next ? [next] : [];
	});
	ctx.run(
		present ? 'Insert page break' : 'Remove page break',
		'view',
		[{ kind: 'parts', sheet: s, parts: ['preserved'] }],
		() => {
			if (result.length) sheet.preserved.set(key, result);
			else sheet.preserved.delete(key);
		},
		{ sheet: s },
	);
}

export function resetPageBreaks(ctx: EditContext, s: number): void {
	const sheet = sheetAt(ctx.workbook, s);
	const next = (['row', 'col'] as const).map((axis) => {
		const key = axis === 'row' ? 'rowBreaks' : 'colBreaks';
		return {
			key,
			xml: (sheet.preserved.get(key) ?? []).flatMap((xml) => {
				const result = patch(xml, undefined, false, axis);
				return result ? [result] : [];
			}),
		};
	});
	ctx.run(
		'Reset page breaks',
		'view',
		[{ kind: 'parts', sheet: s, parts: ['preserved'] }],
		() => {
			for (const entry of next) {
				if (entry.xml.length) sheet.preserved.set(entry.key, entry.xml);
				else sheet.preserved.delete(entry.key);
			}
		},
		{ sheet: s },
	);
}
