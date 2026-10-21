import type { PptxTableCell, PptxTableData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { DEFAULT_TEXT_COLOR } from '../constants';
import { tableCellCss } from './table-cell-css';
import { cellRunStyle } from './table-style';

/**
 * A programmatically inserted table: exactly what `newTableElement` produces
 * (header row + banded rows, no table-style GUID). React rendered these through
 * a path that never consulted the band cascade at all, so they came out flat.
 */
function insertedTable(): PptxTableData {
	return {
		rows: [],
		columnWidths: [0.5, 0.5],
		firstRowHeader: true,
		bandedRows: true,
	} as unknown as PptxTableData;
}

const POS = { rowIndex: 0, cellIndex: 0, rowCount: 4, columnCount: 2 };

describe('tableCellCss', () => {
	it("lays cell text out with PowerPoint's single line spacing and default margins", () => {
		expect(tableCellCss(insertedTable(), undefined, POS)).toMatchObject({
			lineHeight: 1.2,
			paddingLeft: '9.6px',
			paddingRight: '9.6px',
			paddingTop: '4.8px',
			paddingBottom: '4.8px',
		});
	});

	it("keeps a cell's own margins", () => {
		const cell = { text: 'A', style: { marginLeft: 0, marginTop: 2 } };
		expect(tableCellCss(insertedTable(), cell, POS)).toMatchObject({
			paddingLeft: '0px',
			paddingTop: '2px',
			paddingRight: '9.6px',
		});
	});

	it('applies the header emphasis a programmatic table declares', () => {
		const css = tableCellCss(insertedTable(), undefined, POS);
		expect(css.fontWeight).toBe(700);
		expect(css.backgroundColor).toBeDefined();
	});

	it('bands alternate body rows differently', () => {
		const td = insertedTable();
		const band1 = tableCellCss(td, undefined, { ...POS, rowIndex: 1 });
		const band2 = tableCellCss(td, undefined, { ...POS, rowIndex: 2 });
		expect(band1.backgroundColor).not.toBe(band2.backgroundColor);
	});

	it('floors an unstyled cell at the dark slide-text colour', () => {
		// Without the floor the cell inherits the host chrome's `foreground`
		// (#f0efec on the dark preset), i.e. near-white text on a light cell.
		const css = tableCellCss(undefined, undefined, POS);
		expect(css.color).toBe(DEFAULT_TEXT_COLOR);
	});

	it('lets the band layer set the text colour instead of the floor', () => {
		const css = tableCellCss(insertedTable(), undefined, POS);
		expect(css.color).toBe('#ffffff');
	});

	it('lets an explicit cell colour beat both the band and the floor', () => {
		const cell = { text: '', style: { color: '#ff0000' } } as PptxTableCell;
		const css = tableCellCss(insertedTable(), cell, POS);
		expect(css.color).toBe('#ff0000');
	});

	it('keeps a base colour the binding already resolved', () => {
		const css = tableCellCss(undefined, undefined, POS, undefined, { color: '#123456' });
		expect(css.color).toBe('#123456');
	});

	it('layers the explicit cell style over the band fill', () => {
		const cell = { text: '', style: { backgroundColor: '#00ff00' } } as PptxTableCell;
		const css = tableCellCss(insertedTable(), cell, { ...POS, rowIndex: 1 });
		expect(css.backgroundColor).toBe('#00ff00');
	});

	it('lets a cell image fill beat the band fill underneath it', () => {
		// A banded row would otherwise paint a `backgroundColor` here; the image
		// fill must still win, exactly like an explicit solid/gradient fill does.
		const cell = {
			text: '',
			style: { fillMode: 'image', backgroundImageFillData: 'data:image/png;base64,AAAA' },
		} as PptxTableCell;
		const css = tableCellCss(insertedTable(), cell, { ...POS, rowIndex: 1 });
		expect(css.backgroundImage).toBe('url("data:image/png;base64,AAAA")');
	});

	it('renders an explicit zero cell margin as zero padding through the full cascade', () => {
		const cell = { text: '', style: { marginLeft: 0 } } as PptxTableCell;
		const css = tableCellCss(insertedTable(), cell, POS);
		expect(css.paddingLeft).toBe('0px');
	});

	describe('weight variants named as their own family', () => {
		const body = { ...POS, rowIndex: 1 };
		const runWeights = (families: string[], position: typeof POS) => {
			const cell = {
				text: 'x',
				textRuns: families.map((fontFamily) => ({ text: 'x', fontFamily })),
			} as PptxTableCell;
			const css = tableCellCss(insertedTable(), cell, position);
			return {
				cell: css.fontWeight,
				runs: cell.textRuns!.map((run) => cellRunStyle(run, css).fontWeight),
			};
		};

		it('render each regular run at the weight its family names', () => {
			expect(runWeights(['DM Sans Light', 'DM Sans Medium', 'DM Sans Black'], body)).toStrictEqual({
				cell: undefined,
				runs: [300, 500, 900],
			});
			// Whatever the order: a Black run first does not make the cell bold.
			expect(
				runWeights(['DM Sans Black', 'DM Sans Light', 'DM Sans Medium'], body).runs,
			).toStrictEqual([900, 300, 500]);
		});

		it('keep a header row bold when the cell itself is in a Black family', () => {
			// Core copies the first run's typeface onto the cell style.
			const cell = {
				text: 'x',
				style: { fontFamily: 'DM Sans Black' },
				textRuns: ['DM Sans Black', 'DM Sans Light', 'DM Sans Medium'].map((fontFamily) => ({
					text: 'x',
					fontFamily,
				})),
			} as PptxTableCell;
			const header = tableCellCss(insertedTable(), cell, POS);
			expect(header.fontWeight).toBe(900);
			expect(cell.textRuns!.map((run) => cellRunStyle(run, header).fontWeight)).toStrictEqual([
				undefined,
				700,
				700,
			]);
			const row = tableCellCss(insertedTable(), cell, body);
			expect(cell.textRuns!.map((run) => cellRunStyle(run, row).fontWeight)).toStrictEqual([
				undefined,
				300,
				500,
			]);
		});

		it('keep a header row bold, and Black at 900', () => {
			expect(runWeights(['DM Sans Light', 'DM Sans Medium', 'DM Sans Black'], POS)).toStrictEqual({
				cell: 700,
				runs: [undefined, undefined, 900],
			});
		});

		it('render a cell and its runs in a style-named family italic', () => {
			const cell = {
				text: 'x',
				style: { fontFamily: 'Lato Light Italic' },
				textRuns: [
					{ text: 'x', fontFamily: 'Lato Light Italic' },
					{ text: 'y', fontFamily: 'Lato' },
				],
			} as PptxTableCell;
			const css = tableCellCss(insertedTable(), cell, body);
			expect(css.fontStyle).toBe('italic');
			expect(css.fontWeight).toBe(300);
			expect(cellRunStyle(cell.textRuns![0]!, css).fontStyle).toBe('italic');
		});

		it('use the substitution chain for a cell family', () => {
			const cell = { text: 'A', style: { fontFamily: 'DM Sans Medium' } } as PptxTableCell;
			const css = tableCellCss(insertedTable(), cell, body);
			expect(css.fontFamily).toMatch(/^"DM Sans Medium", "DM Sans"/u);
			expect(css.fontWeight).toBe(500);
		});
	});
});
