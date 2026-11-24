import { describe, expect, it } from 'vitest';
import type { PdfTextSpan } from '../../pdf/index';
import { pageLines } from './lines';
import { cellParagraphs, joinTables, type TableBlock } from './tables';

const span = (x: number, y: number, text: string): PdfTextSpan => ({
	text,
	x,
	y,
	width: text.length * 5,
	fontSize: 11,
	font: {
		name: 'Calibri',
		family: 'Calibri',
		bold: false,
		italic: false,
		serif: false,
		monospace: false,
	},
	color: '#000000',
	angle: 0,
	rise: 0,
});

const WIDTHS = [100, 300];
const BORDER = { width: 0.5, color: '#000000' };

/** A two-column ruled table from rows of cell texts (`\n` between a cell's lines). */
function table(rows: string[][]): TableBlock {
	return {
		kind: 'table',
		page: 0,
		top: 100,
		bottom: 100 + 20 * rows.length,
		left: 72,
		columns: WIDTHS,
		rows: rows.map(() => ({ height: 20 })),
		padding: 5.4,
		cells: rows.flatMap((texts, row) =>
			texts.map((text, col) => {
				const x = 72 + (col ? WIDTHS[0]! : 0);
				const lines = pageLines(
					0,
					text
						.split('\n')
						.filter(Boolean)
						.map((line, at) => span(x + 5.4, 100 + row * 20 + at * 14, line)),
				).lines;
				return {
					row,
					col,
					rowSpan: 1,
					colSpan: 1,
					borders: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
					blocks: cellParagraphs(lines, x + 5.4, x + WIDTHS[col]! - 5.4, 11),
				};
			}),
		),
	};
}

const texts = (block: TableBlock) =>
	Array.from({ length: block.rows.length }, (_, row) =>
		block.cells
			.filter((cell) => cell.row === row)
			.map((cell) =>
				cell.blocks.map((item) => item.fragments.map((f) => f.text).join('')).join('/'),
			),
	);

describe('joinTables and a row cut by the page break', () => {
	it('joins a continuation that leaves a never-empty column empty', () => {
		const first = table([
			['Réf.', 'Texte'],
			['R1', 'Un.'],
			['R2', 'Début du texte.'],
		]);
		const next = table([
			['Réf.', 'Texte'],
			['', 'Suite du texte.'],
			['R3', 'Trois.'],
		]);
		expect(joinTables(first, next)).toBe(true);
		expect(texts(first)).toEqual([
			['Réf.', 'Texte'],
			['R1', 'Un.'],
			['R2', 'Début du texte./Suite du texte.'],
			['R3', 'Trois.'],
		]);
		expect(first.rows[0]!.header).toBe(true);
		expect(first.rows[2]!.height).toBe(40);
	});

	it('joins cells cut mid-sentence into one paragraph', () => {
		const first = table([
			['A', 'Un.'],
			['B', 'Deux.'],
			['Le début', 'Une phrase coupée par'],
		]);
		const next = table([
			['de la cellule', 'le saut de page.'],
			['C', 'Trois.'],
		]);
		expect(joinTables(first, next)).toBe(true);
		expect(texts(first)[2]).toEqual([
			'Le début de la cellule',
			'Une phrase coupée par le saut de page.',
		]);
		expect(first.rows).toHaveLength(4);
	});

	it('keeps rows apart when the empty cell is usual in the table or the text starts anew', () => {
		const grouped = table([
			['Groupe', 'Article 1'],
			['', 'Article 2'],
			['', 'Article 3'],
		]);
		expect(
			joinTables(
				grouped,
				table([
					['', 'Article 4'],
					['', 'Article 5'],
				]),
			),
		).toBe(true);
		expect(grouped.rows).toHaveLength(5);
		const plain = table([
			['R1', 'Un'],
			['R2', 'Deux'],
		]);
		expect(
			joinTables(
				plain,
				table([
					['R3', 'Trois'],
					['R4', 'Quatre'],
				]),
			),
		).toBe(true);
		expect(plain.rows).toHaveLength(4);
	});
});
