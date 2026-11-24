import { describe, expect, it } from 'vitest';
import type { PdfTextSpan } from '../../pdf/index';
import { unruledTables } from './borderless';
import { applyLeaders } from './leaders';
import { pageLines, type Line } from './lines';

const span = (x: number, y: number, text: string): PdfTextSpan => ({
	text,
	x,
	y,
	// About 5 points a character.
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

/** Lines from `[baseline, [x, text]...]` rows, read like a page. */
function page(...rows: [number, ...[number, string][]][]): Line[] {
	return pageLines(
		0,
		rows.flatMap(([y, ...cells]) => cells.map(([x, text]) => span(x, y, text))),
	).lines;
}

const cells = (lines: Line[]) => {
	const { tables } = unruledTables(0, lines, 11);
	return tables.map((table) =>
		Array.from({ length: table.rows.length }, (_, row) =>
			table.cells
				.filter((cell) => cell.row === row)
				.map((cell) =>
					cell.blocks
						.map((block) => block.fragments.map((fragment) => fragment.text).join(''))
						.join('/'),
				),
		),
	);
};

describe('unruledTables', () => {
	it('keeps wrapped cells in their row and leaves empty cells empty', () => {
		const lines = page(
			[100, [72, 'Nom'], [172, 'Description'], [372, 'Note']],
			[114, [72, 'Alpha'], [172, 'Ligne longue qui remplit la colonne'], [372, 'Oui']],
			[128, [172, 'jusqu’au bord.']],
			[142, [72, 'Beta'], [372, 'Non']],
			[156, [72, 'Gamma'], [172, 'Court']],
			[186, [72, 'Le paragraphe qui suit le tableau est une ligne de texte ordinaire et longue.']],
		);
		expect(cells(lines)).toEqual([
			[
				['Nom', 'Description', 'Note'],
				['Alpha', 'Ligne longue qui remplit la colonne jusqu’au bord.', 'Oui'],
				['Beta', '', 'Non'],
				['Gamma', 'Court', ''],
			],
		]);
		expect(unruledTables(0, lines, 11).lines.map((line) => line.text)).toEqual([
			'Le paragraphe qui suit le tableau est une ligne de texte ordinaire et longue.',
		]);
	});

	it('starts a column for an empty top-left header cell', () => {
		const lines = page(
			[100, [172, 'Prix'], [272, 'Stock']],
			[114, [72, 'Pommes'], [172, '2'], [272, '40']],
			[128, [72, 'Poires'], [172, '3'], [272, '25']],
		);
		expect(cells(lines)).toEqual([
			[
				['', 'Prix', 'Stock'],
				['Pommes', '2', '40'],
				['Poires', '3', '25'],
			],
		]);
	});

	it('leaves paragraphs, tab-separated lists and leader lines alone', () => {
		const prose = page(
			[100, [72, 'Un paragraphe ordinaire sur plusieurs lignes de texte, sans tabulation.']],
			[114, [72, 'La deuxième ligne du paragraphe continue la phrase.']],
			[128, [72, 'La troisième ligne la termine.']],
		);
		expect(cells(prose)).toEqual([]);
		const list = page(
			[100, [72, '1.'], [108, 'Premier élément de la liste.']],
			[114, [72, '2.'], [108, 'Deuxième élément, plus long, qui revient']],
			[128, [108, 'à la ligne.']],
			[142, [72, '3.'], [108, 'Troisième élément.']],
		);
		expect(cells(list)).toEqual([]);
		const toc = page(
			[100, [72, 'Introduction'], [132, '..............................'], [300, '3']],
			[114, [72, 'Méthode'], [107, '...................................'], [300, '5']],
			[128, [72, 'Résultats'], [117, '.................................'], [300, '8']],
		);
		applyLeaders(toc);
		expect(cells(toc)).toEqual([]);
	});
});
