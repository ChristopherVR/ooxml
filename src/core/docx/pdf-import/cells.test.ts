import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import type { Block, Paragraph, Table } from '../model';
import { loadDocx } from '../parse';
import { importPdf } from './index';

const fixture = async (name: string) =>
	new Uint8Array(await readFile(new URL(`./__fixtures__/${name}`, import.meta.url)));
const textOf = (paragraph: Paragraph) => paragraph.runs.map((run) => run.text).join('');
const tables = (blocks: Block[]) =>
	blocks.filter((block): block is Table => block.type === 'table');
const paragraphs = (blocks: Block[]) =>
	blocks.filter((block): block is Paragraph => block.type === 'paragraph');
const grid = (table: Table) =>
	table.rows.map((row) => row.map((cell) => cell.paragraphs.map(textOf).join('\n')));

describe('table cells (cellules.pdf, exported by Word 16)', async () => {
	const { model, document } = await importPdf(await fixture('cellules.pdf'));
	const [plain, lists, split] = tables(model.blocks);

	it('rebuilds a borderless table with wrapped and empty cells', () => {
		expect(tables(model.blocks)).toHaveLength(3);
		expect(grid(plain!)).toEqual([
			['Élément', 'Description', 'Remarque'],
			[
				'Papier',
				'Feuilles blanches au format A4, achetées par ramettes de cinq cents feuilles.',
				'Stock suffisant',
			],
			['Encre', '', 'À commander'],
			['Agrafes', 'Petite boîte', ''],
			['', 'Ligne sans élément', 'Note libre'],
			[
				'Classeurs',
				'Classeurs à levier, dos de huit centimètres, en carton recouvert.',
				'Livrés en janvier, rangés dans l’armoire du couloir.',
			],
		]);
		expect(plain!.rows[0]![0]!.borders?.top).toBeUndefined();
	});

	it('makes the bullets and numbers in cells real list items', () => {
		expect(grid(lists!)[1]).toEqual([
			'Préparer',
			'Réunir les pièces.\nVérifier les dates.\nClasser le dossier.',
			'Relire la liste.\nSigner le registre.',
		]);
		const actions = lists!.rows[1]![1]!.paragraphs;
		const checks = lists!.rows[1]![2]!.paragraphs;
		const later = lists!.rows[2]![2]!.paragraphs;
		expect(actions.every((paragraph) => paragraph.numbering?.level === 0)).toBe(true);
		expect(checks.every((paragraph) => paragraph.numbering)).toBe(true);
		const catalog = model.numberingCatalog!;
		const level = (numId: number) =>
			catalog.abstractNums[catalog.nums[String(numId)]!.abstractNumId]!.levels[0]!;
		expect(level(actions[0]!.numbering!.numId).numFmt).toBe('bullet');
		expect(level(checks[0]!.numbering!.numId)).toMatchObject({ numFmt: 'decimal', lvlText: '%1.' });
		// The numbers go on in the next row (3. to 5.): the same list.
		expect(later.map((paragraph) => paragraph.numbering?.numId)).toEqual(
			later.map(() => checks[0]!.numbering!.numId),
		);
		expect(
			grid(lists!)
				.flat()
				.some((text) => /^(•|\d\.)/m.test(text)),
		).toBe(false);
	});

	it('joins the row the page break cut into one row', async () => {
		expect(grid(split!).map((row) => row[0])).toEqual([
			'Réf.',
			'R1',
			'R2',
			'R3',
			'R4',
			'R5',
			'R6',
			'R7',
		]);
		expect(split!.rowProperties?.[0]?.header).toBe(true);
		const cut = grid(split!)[5]!;
		expect(cut[2]).toBe('50');
		expect(cut[1]!.split('\n').map((text) => text.split(' ')[0])).toEqual([
			'Première',
			'Deuxième',
			'Troisième',
			'Quatrième',
			'Cinquième',
			'Sixième',
		]);
		const reopened = await loadDocx(await document.save());
		expect(tables(reopened.model.blocks).map((table) => table.rows.length)).toEqual([6, 3, 8]);
	});

	it('leaves the paragraphs around the tables alone', () => {
		const text = paragraphs(model.blocks).map(textOf);
		expect(text).toContain('Un tableau avec des listes dans ses cellules :');
		expect(text[text.length - 1]).toBe('Fin des tableaux.');
		expect(text.some((line) => /Papier|Réunir|Article/.test(line))).toBe(false);
	});
});
