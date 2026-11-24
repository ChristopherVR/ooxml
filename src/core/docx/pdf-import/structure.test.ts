import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import type { Block, Paragraph, Table } from '../model';
import { loadDocx } from '../parse';
import { importPdf } from './index';

const fixture = async (name: string) =>
	new Uint8Array(await readFile(new URL(`./__fixtures__/${name}`, import.meta.url)));
const textOf = (paragraph: Paragraph) => paragraph.runs.map((run) => run.text).join('');
const paragraphs = (blocks: Block[]) =>
	blocks.filter((block): block is Paragraph => block.type === 'paragraph');
const tables = (blocks: Block[]) =>
	blocks.filter((block): block is Table => block.type === 'table');
const cellText = (table: Table, row: number, cell: number) =>
	table.rows[row]![cell]!.paragraphs.map(textOf).join('\n');

describe('tables (tableaux.pdf, exported by Word 16)', async () => {
	const { model, document } = await importPdf(await fixture('tableaux.pdf'));
	const [ruled, plain] = tables(model.blocks);

	it('rebuilds the ruled table across the page break as one table', () => {
		expect(tables(model.blocks)).toHaveLength(2);
		expect(ruled!.grid).toHaveLength(4);
		// The header row, the four rows with merged cells and 34 article rows; the header the PDF
		// repeats on page 2 is not a row again.
		expect(ruled!.rows).toHaveLength(38);
		expect(ruled!.rows.map((_, row) => cellText(ruled!, row, 0))).toEqual([
			'Code',
			'A0',
			'B0',
			'B1',
			...Array.from({ length: 34 }, (_, index) => `C${index + 1}`),
		]);
		expect(ruled!.rowProperties?.[0]?.header).toBe(true);
		expect(ruled!.rows[37]!.map((cell) => cell.paragraphs.map(textOf).join(''))).toEqual([
			'C34',
			'Article 34',
			'Est',
			'331',
		]);
	});

	it('keeps merged cells, shading and borders', () => {
		const header = ruled!.rows[0]!;
		expect(header.map((cell) => cell.paragraphs.map(textOf).join(''))).toEqual([
			'Code',
			'Produit',
			'Région',
			'Montant',
		]);
		expect(header.every((cell) => cell.shadingFill?.toUpperCase().endsWith('D9E2F3'))).toBe(true);
		expect(ruled!.rows[1]!).toHaveLength(3);
		expect(ruled!.rows[1]![1]).toMatchObject({ gridSpan: 2 });
		expect(cellText(ruled!, 1, 1)).toBe('Synthèse sur deux colonnes');
		expect(ruled!.rows[2]![1]).toMatchObject({ verticalMerge: 'restart' });
		expect(cellText(ruled!, 2, 1)).toBe('Groupe fusionné');
		expect(ruled!.rows[3]![1]).toMatchObject({ verticalMerge: 'continue' });
		expect(ruled!.rows[4]![0]!.borders?.top).toBeTruthy();
		expect(ruled!.rows[4]![0]!.paragraphs.flatMap((p) => p.runs).some((run) => run.underline)).toBe(
			false,
		);
	});

	it('rebuilds the borderless aligned table', () => {
		expect(
			plain!.rows.map((row) => row.map((cell) => cell.paragraphs.map(textOf).join(''))),
		).toEqual([
			['Nom', 'Ville', 'Pays'],
			['Alice', 'Lyon', 'France'],
			['Bruno', 'Genève', 'Suisse'],
			['Chloé', 'Namur', 'Belgique'],
		]);
		expect(plain!.rows[0]![0]!.borders?.top).toBeUndefined();
	});

	it('leaves table text out of the paragraphs around it', async () => {
		const text = paragraphs(model.blocks).map(textOf);
		expect(text.some((line) => /Article|Montant|Alice|\t/.test(line))).toBe(false);
		expect(text).toContain('Le tableau suivant n’a pas de bordures.');
		expect(text[text.length - 1]).toBe('Fin des tableaux.');
		const reopened = await loadDocx(await document.save());
		expect(tables(reopened.model.blocks).map((table) => table.rows.length)).toEqual([38, 4]);
	});
});

describe('tables in another producer’s PDF (lettre.pdf, from exportDocxPdf)', () => {
	it('rebuilds the head, body and foot tables and keeps the margins', async () => {
		const { model } = await importPdf(await fixture('lettre.pdf'));
		expect(
			tables(model.blocks).map((table) => [
				table.rows.length,
				table.grid?.length,
				cellText(table, 0, 0),
			]),
		).toEqual([
			[2, 2, 'Notre référence : DOS-2026-001'],
			[2, 2, 'Objet'],
			[2, 3, 'Cabinet Exemple'],
		]);
		expect(paragraphs(model.blocks).map(textOf).filter(Boolean)).toEqual([
			'Corps de la lettre.',
			'Fin du document.',
		]);
		const section = model.sections![0]!;
		expect(Math.abs(section.marginRightTwips - section.marginLeftTwips)).toBeLessThan(40);
	});
});

describe('lists (listes.pdf, exported by Word 16)', async () => {
	const { model } = await importPdf(await fixture('listes.pdf'));
	const body = paragraphs(model.blocks);
	const items = body.filter((paragraph) => paragraph.numbering);

	it('makes bullets and numbers real list items, the typed marker removed', () => {
		expect(items.map((item) => [textOf(item).slice(0, 14), item.numbering!.level])).toEqual([
			['Premier point ', 0],
			['Deuxième point', 0],
			['Sous-point un.', 1],
			['Sous-point deu', 1],
			['Troisième poin', 0],
			['Ouvrir le docu', 0],
			['Lire chaque se', 0],
			['Noter les rema', 1],
			['Classer les re', 1],
			['Enregistrer le', 0],
		]);
		const [bullets, numbers] = [items[0]!.numbering!.numId, items[5]!.numbering!.numId];
		expect(bullets).not.toBe(numbers);
		expect(items.slice(0, 5).every((item) => item.numbering!.numId === bullets)).toBe(true);
		expect(items.slice(5).every((item) => item.numbering!.numId === numbers)).toBe(true);
	});

	it('keeps a wrapped item as one paragraph and writes the levels', () => {
		expect(textOf(items[1]!)).toMatch(/^Deuxième point.*entreprise réelle\.$/);
		expect(body.map(textOf)).toContain('Une liste numérotée :');
		const catalog = model.numberingCatalog!;
		const levels = (numId: number) =>
			catalog.abstractNums[catalog.nums[String(numId)]!.abstractNumId]!.levels;
		const bullet = levels(items[0]!.numbering!.numId);
		expect([bullet[0]!.numFmt, bullet[0]!.markerFormat?.fontFamily]).toEqual(['bullet', 'Symbol']);
		expect([bullet[1]!.lvlText, bullet[1]!.markerFormat?.fontFamily]).toEqual(['o', 'Courier New']);
		const numbered = levels(items[5]!.numbering!.numId);
		expect([numbered[0]!.numFmt, numbered[0]!.lvlText]).toEqual(['decimal', '%1.']);
		expect([numbered[1]!.numFmt, numbered[1]!.lvlText]).toEqual(['lowerLetter', '%2)']);
		expect(numbered[0]!.indentLeftTwips).toBe(720);
		expect(numbered[0]!.hangingTwips).toBe(360);
	});
});

describe('columns (colonnes.pdf, exported by Word 16)', async () => {
	const { model } = await importPdf(await fixture('colonnes.pdf'));
	const body = paragraphs(model.blocks).map(textOf);

	it('reads the two-column region column by column', () => {
		const starts = body.map((text) => text.split(' ')[0]);
		expect(starts.slice(2)).toEqual([
			'Le',
			'Premier',
			'Deuxième',
			'Troisième',
			'Quatrième',
			'Cinquième',
			'Sixième',
			'Retour',
			'Page',
		]);
		expect(body.some((text) => text.includes('\t'))).toBe(false);
	});

	it('puts it in a continuous two-column section between one-column ones', () => {
		const sections = model.sections!;
		expect(sections.map((section) => section.columns.count)).toEqual([1, 2, 1]);
		expect(sections[1]!.type).toBe('continuous');
		expect(sections[2]!.type).toBe('continuous');
		expect(Math.abs((sections[1]!.columns.spacingTwips ?? 0) - 708)).toBeLessThan(80);
	});
});
