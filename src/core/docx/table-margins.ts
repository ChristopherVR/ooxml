// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Table } from './model';
import type { TableCellMargins, TableStyleCatalog, TableStyleDefinition } from './table-model';
import { first, getW, type XmlElement } from './xml';
import { parseSignedTwips } from './simple-types';
import { signedTwips, type SignedTwips } from './units';

/** Cell margins from `w:tcMar` or `w:tblCellMar`, in twips (`start`/`end` read as left/right). */
export function parseCellMargins(element: XmlElement | undefined): TableCellMargins | undefined {
	if (!element) return undefined;
	const parsed: TableCellMargins = {};
	for (const [side, logical] of [
		['top', 'top'],
		['bottom', 'bottom'],
		['left', 'start'],
		['right', 'end'],
	] as const) {
		const edge = first(element, logical) ?? first(element, side);
		const type = getW(edge, 'type') ?? 'dxa';
		const value =
			type === 'nil'
				? parseSignedTwips('0')
				: type === 'dxa'
					? parseSignedTwips(getW(edge, 'w'))
					: undefined;
		if (value !== undefined) parsed[side] = value;
	}
	return Object.keys(parsed).length ? parsed : undefined;
}

const SIDES = ['top', 'right', 'bottom', 'left'] as const;
type Margins = Record<(typeof SIDES)[number], SignedTwips>;

/** Word's built-in Normal Table margins, for a model that carries no styles part at all. */
const BUILT_IN: Margins = {
	top: signedTwips(0),
	right: signedTwips(108),
	bottom: signedTwips(0),
	left: signedTwips(108),
};

/**
 * What Word uses for a side no style sets, when the styles part has no table style to give it:
 * 10 twips left and right (measured: Word lays the table out so and writes them on save).
 */
const UNSTYLED: Margins = {
	top: signedTwips(0),
	right: signedTwips(10),
	bottom: signedTwips(0),
	left: signedTwips(10),
};

/** A style and the styles it is based on, nearest first (a cycle stops the walk). */
function styleChain(id: string | undefined, catalog: TableStyleCatalog): TableStyleDefinition[] {
	const chain: TableStyleDefinition[] = [];
	let style = id ? catalog.styles[id] : undefined;
	while (style && !chain.includes(style)) {
		chain.push(style);
		style = style.basedOn ? catalog.styles[style.basedOn] : undefined;
	}
	return chain;
}

/**
 * A table's default cell margins, side by side: its own `w:tblCellMar`, then its table style and
 * the styles that one is based on, then the document's default table style and its chain
 * (measured, `__fixtures__/table-autofit/style-*`). A side none of them sets takes Word's fallback
 * (`UNSTYLED`). A model without a styles part keeps Word's built-in Normal Table margins (108 twips left
 * and right), as the editor's new documents do.
 */
export function resolveTableCellMargins(
	table: Pick<Table, 'cellMargins' | 'style'>,
	catalog: TableStyleCatalog | undefined,
): Margins {
	if (!catalog) return { ...BUILT_IN, ...table.cellMargins };
	const defaultStyle = Object.values(catalog.styles).find((style) => style.isDefault);
	const layers = [
		table.cellMargins,
		...styleChain(table.style, catalog).map((style) => style.cellMargins),
		...styleChain(defaultStyle?.id, catalog).map((style) => style.cellMargins),
	];
	const result: Margins = { ...UNSTYLED };
	for (const side of SIDES) {
		const value = layers.find((layer) => layer?.[side] !== undefined)?.[side];
		if (value !== undefined) result[side] = value;
	}
	return result;
}
