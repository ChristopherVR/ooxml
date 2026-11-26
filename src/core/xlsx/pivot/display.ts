import { NS, elements, first, type XmlElement } from '../../xml/index';
import type { PivotDisplayOptions, PivotStyleOptions } from '../model';

const styleFlags = [
	'showRowHeaders',
	'showColHeaders',
	'showRowStripes',
	'showColStripes',
] as const;
// Excel's RowGrand controls the grand-total column (XML colGrandTotals).
// ColumnGrand controls the grand-total row (XML rowGrandTotals).
const tableFlags = [
	['rowGrandTotals', 'colGrandTotals'],
	['colGrandTotals', 'rowGrandTotals'],
	['showHeaders', 'showHeaders'],
	['showDrill', 'showDrill'],
] as const;

export function pivotFlag(node: XmlElement | undefined, name: string): boolean | undefined {
	const value = node?.getAttribute(name);
	if (value == null) return;
	if (value === '1' || value === 'true') return true;
	if (value === '0' || value === 'false') return false;
	throw new Error('Invalid PivotTable display option.');
}

export function readPivotDisplay(
	root: XmlElement,
	fields: XmlElement[],
	rows: XmlElement[],
	valuesOnRows: boolean,
): PivotDisplayOptions | undefined {
	const display: PivotDisplayOptions = {};
	for (const [flag, attr] of tableFlags) if (pivotFlag(root, attr) === false) display[flag] = false;
	const styleNode = first(root, 'pivotTableStyleInfo', NS.x);
	if (styleNode) {
		const style: PivotStyleOptions = { name: styleNode.getAttribute('name') ?? '' };
		for (const flag of styleFlags) {
			const value = pivotFlag(styleNode, flag);
			if (value !== undefined) style[flag] = value;
		}
		display.style = style;
	}
	const blank = fields.filter((field) => pivotFlag(field, 'insertBlankRow'));
	if (blank.length) {
		if (
			blank.length !== rows.length ||
			blank.some((field) => !rows.includes(field)) ||
			valuesOnRows
		)
			throw new Error('Mixed PivotTable blank rows are preserved.');
		display.blankRows = true;
	}
	const items = (field?: XmlElement) => {
		const list = first(field, 'items', NS.x);
		return list ? elements(list).filter((item) => item.hasAttribute('x')) : [];
	};
	const collapsed = fields.flatMap((field) =>
		items(field).filter((item) => pivotFlag(item, 'sd') === false),
	);
	if (collapsed.length) {
		if (
			rows.length < 2 ||
			valuesOnRows ||
			collapsed.length !== items(rows[0]).length ||
			items(rows[0]).some((item) => pivotFlag(item, 'sd') !== false)
		)
			throw new Error('Mixed PivotTable item expansion is preserved.');
		display.collapsed = true;
	}
	return Object.keys(display).length ? display : undefined;
}

export function writePivotDisplay(root: XmlElement, display: PivotDisplayOptions | undefined) {
	for (const [flag, attr] of tableFlags) {
		const value = display?.[flag];
		if (value !== undefined || root.hasAttribute(attr))
			root.setAttribute(attr, value === false ? '0' : '1');
	}
	if (!display?.style) return;
	let style = first(root, 'pivotTableStyleInfo', NS.x);
	if (!style) {
		style = root.ownerDocument!.createElementNS(NS.x, 'pivotTableStyleInfo');
		root.insertBefore(style, first(root, 'extLst', NS.x) ?? null);
	}
	if (display.style.name === '') style.removeAttribute('name');
	else if (display.style.name !== undefined) style.setAttribute('name', display.style.name);
	for (const flag of styleFlags) {
		const value = display.style[flag];
		if (value !== undefined) style.setAttribute(flag, value ? '1' : '0');
	}
}
