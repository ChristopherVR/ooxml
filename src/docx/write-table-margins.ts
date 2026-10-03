import type { TableCellMargins } from './table-model.js';
import { orderChildren } from './element-order.js';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

const SIDES = ['top', 'left', 'bottom', 'right'] as const;
const ORDER = ['top', 'start', 'left', 'bottom', 'end', 'right'];

/** Patch only changed margin sides, retaining logical edge names and unrelated XML. */
export function patchMarginProperties(
	doc: XmlDocument,
	host: XmlElement,
	propertiesTag: 'tblPr' | 'tcPr',
	marginTag: 'tblCellMar' | 'tcMar',
	propertiesOrder: readonly string[],
	next: TableCellMargins | undefined,
	base: TableCellMargins | undefined,
): void {
	const changed = SIDES.filter((side) => next?.[side] !== base?.[side]);
	if (!changed.length) return;
	for (const side of changed) {
		const value = next?.[side];
		if (value !== undefined && (!Number.isSafeInteger(value) || value < 0))
			throw new Error('Table cell margins must be nonnegative whole twips.');
	}
	let properties = first(host, propertiesTag);
	if (!properties) {
		properties = makeW(doc, propertiesTag);
		host.insertBefore(properties, host.firstChild);
	}
	let margins = first(properties, marginTag);
	if (!margins) {
		margins = makeW(doc, marginTag);
		properties.appendChild(margins);
	}
	for (const side of changed) {
		const logical = side === 'left' ? 'start' : side === 'right' ? 'end' : side;
		const edge = first(margins, logical) ?? first(margins, side) ?? makeW(doc, side);
		for (const alias of new Set([side, logical]))
			for (const old of children(margins, alias)) margins.removeChild(old);
		const value = next?.[side];
		if (value === undefined) continue;
		edge.setAttributeNS(WORD_NS, 'w:w', String(value));
		edge.setAttributeNS(WORD_NS, 'w:type', 'dxa');
		margins.appendChild(edge);
	}
	if (
		!Array.from(margins.childNodes).some((node) => node.nodeType === 1) &&
		!margins.attributes.length
	)
		properties.removeChild(margins);
	else orderChildren(margins, ORDER);
	orderChildren(properties, propertiesOrder);
	if (
		!Array.from(properties.childNodes).some((node) => node.nodeType === 1) &&
		!properties.attributes.length
	)
		host.removeChild(properties);
}
