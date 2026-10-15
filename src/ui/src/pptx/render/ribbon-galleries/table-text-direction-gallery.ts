/**
 * Table Layout > Alignment > Text Direction: Horizontal, Rotate all text 90, Rotate all text 270
 * and Stacked, written as `a:tcPr/@vert` (`vert`, `vert270`, `wordArtVert`) on the selected
 * cells, or on every cell when the table frame is selected.
 *
 * @module render/ribbon-galleries/table-text-direction-gallery
 */
import type { PptxTableCellStyle } from 'ooxml-core/pptx';

import type { RibbonGalleryModule } from './gallery-module';
import type { RibbonGalleryContext } from './gallery-types';
import { styleCells } from './table-layout-ops';

type Direction = NonNullable<PptxTableCellStyle['textDirection']>;

interface Entry {
	id: string;
	label: string;
	/** `a:tcPr/@vert`; undefined is horizontal. */
	vert: Direction | undefined;
	transform: string;
}

const ENTRIES: readonly Entry[] = [
	{ id: 'horizontal', label: 'Horizontal', vert: undefined, transform: '' },
	{ id: 'rotate90', label: 'Rotate all text 90°', vert: 'vert', transform: 'rotate(90 18 15)' },
	{
		id: 'rotate270',
		label: 'Rotate all text 270°',
		vert: 'vert270',
		transform: 'rotate(-90 18 15)',
	},
	{ id: 'stacked', label: 'Stacked', vert: 'wordArtVert', transform: '' },
];

/** A fixed glyph tile; built from the catalogue only, never from user text. */
function tile(entry: Entry): string {
	const body =
		entry.id === 'stacked'
			? '<text x="18" y="9" font-size="9" text-anchor="middle" fill="currentColor">A</text><text x="18" y="18" font-size="9" text-anchor="middle" fill="currentColor">b</text><text x="18" y="27" font-size="9" text-anchor="middle" fill="currentColor">c</text>'
			: `<text x="18" y="19" font-size="11" text-anchor="middle" fill="currentColor"${entry.transform ? ` transform="${entry.transform}"` : ''}>Abc</text>`;
	return `<svg xmlns="http://www.w3.org/2000/svg" width="36" height="30" viewBox="0 0 36 30" role="img" aria-hidden="true">${body}</svg>`;
}

function currentDirection(ctx: RibbonGalleryContext): Direction | undefined {
	const element = ctx.element;
	if (element?.type !== 'table' || !element.tableData) {
		return undefined;
	}
	const cells = ctx.selectedTableCells ?? [];
	const first = cells[0];
	const cell = first
		? element.tableData.rows[first.row]?.cells[first.column]
		: element.tableData.rows[0]?.cells[0];
	return cell?.style?.textDirection;
}

export const TABLE_TEXT_DIRECTION_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const current = currentDirection(ctx);
		return {
			id: 'tableTextDirection',
			labelKey: 'pptx.gallery.tableLayout.textDirection',
			label: 'Text Direction',
			disabled: ctx.element?.type !== 'table',
			sections: [
				{
					id: 'direction',
					columns: 4,
					tileWidth: 36,
					tileHeight: 30,
					items: ENTRIES.map((entry) => ({
						id: entry.id,
						labelKey: `pptx.gallery.tableLayout.direction.${entry.id}`,
						label: entry.label,
						previewSvg: tile(entry),
						applied: entry.vert === current,
					})),
				},
			],
		};
	},
	apply(itemId, ctx) {
		const element = ctx.element;
		const entry = ENTRIES.find((candidate) => candidate.id === itemId);
		if (element?.type !== 'table' || !entry) {
			return null;
		}
		const cells = ctx.selectedTableCells ?? [];
		return entry.vert
			? styleCells(element, cells, { textDirection: entry.vert })
			: styleCells(element, cells, {}, ['textDirection']);
	},
};
