import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../parse';
import type { Block, Paragraph } from '../model';
import { buildPdf, simplePdf, stream } from '../../pdf/test-support/build';
import { importPdf } from './index';

const fixture = async (name: string) =>
	new Uint8Array(await readFile(new URL(`./__fixtures__/${name}`, import.meta.url)));
const paragraphs = (blocks: Block[]) =>
	blocks.filter((block): block is Paragraph => block.type === 'paragraph');
const textOf = (paragraph: Paragraph) => paragraph.runs.map((run) => run.text).join('');

describe('importPdf on Word footnotes, links and headings (notes.pdf, exported by Word 16)', async () => {
	const { model, document } = await importPdf(await fixture('notes.pdf'));
	const body = paragraphs(model.blocks);
	const notes = model.footnotes ?? [];

	it('makes the notes at the page bottoms real footnotes, referenced from the body', () => {
		expect(notes.map((note) => paragraphs(note.blocks).map(textOf).join('').trim())).toEqual([
			'Une première note de bas de page, courte.',
			'Une deuxième note, plus longue, qui explique un point de détail du paragraphe sans encombrer le texte principal.',
			'La note de la deuxième page.',
		]);
		const references = body.flatMap((paragraph) =>
			paragraph.runs.filter((run) => run.noteReference).map((run) => run.noteReference!.id),
		);
		expect(references).toEqual(notes.map((note) => note.id));
		const text = body.map(textOf).join(' ');
		expect(text).not.toContain('Une première note');
		expect(text).not.toMatch(/imprimée\.1/);
	});

	it('takes heading levels from the outline and the structure tags', () => {
		const headings = body
			.filter((paragraph) => /^Heading[1-6]$/.test(paragraph.style ?? ''))
			.map((paragraph) => [paragraph.style, textOf(paragraph)]);
		expect(headings).toEqual(
			expect.arrayContaining([
				['Heading1', 'Introduction'],
				['Heading2', 'Contexte'],
				['Heading1', 'Méthode'],
				['Heading2', 'Étapes'],
				['Heading1', 'Annexe'],
			]),
		);
	});

	it('turns internal links into links to bookmarks at their target paragraphs', () => {
		const anchorOf = (text: string) =>
			body.flatMap((paragraph) => paragraph.runs).find((run) => run.text.includes(text))?.link
				?.anchor;
		const bookmarkOf = (text: string) =>
			body.find((paragraph) => textOf(paragraph) === text)?.bookmarks?.[0];
		expect(anchorOf('voir l’annexe')).toBeDefined();
		expect(anchorOf('voir l’annexe')).toBe(bookmarkOf('Annexe'));
		// The table of contents links to each heading, with its dot leaders as a tab stop.
		const toc = body.filter((paragraph) => /\t\d$/.test(textOf(paragraph)));
		expect(toc.map(textOf)).toEqual([
			'Introduction\t1',
			'Contexte\t1',
			'Méthode\t1',
			'Étapes\t2',
			'Annexe\t2',
		]);
		expect(toc.map((paragraph) => paragraph.runs[0]!.link?.anchor)).toEqual(
			['Introduction', 'Contexte', 'Méthode', 'Étapes', 'Annexe'].map(bookmarkOf),
		);
		expect(toc[0]!.tabStops).toEqual([expect.objectContaining({ align: 'right', leader: 'dot' })]);
	});

	it('saves a document that keeps the footnotes and bookmarks', async () => {
		const reopened = await loadDocx(await document.save());
		expect(reopened.model.footnotes).toHaveLength(3);
		expect(paragraphs(reopened.model.blocks).some((paragraph) => paragraph.bookmarks?.length)).toBe(
			true,
		);
	});
});

describe('importPdf on rotated text (rotation.pdf, exported by Word 16)', () => {
	it('puts text running up or down in vertical text boxes', async () => {
		const { model, report } = await importPdf(await fixture('rotation.pdf'));
		const boxes = paragraphs(model.blocks)
			.flatMap((paragraph) => paragraph.runs)
			.flatMap((run) => (run.image?.textBoxContent ? [run.image.textBoxContent] : []));
		expect(
			boxes.map((box) => [box.vertical, paragraphs(box.blocks).map(textOf).join('').trim()]),
		).toEqual(
			expect.arrayContaining([
				['vert270', 'Écrit de bas en haut'],
				['vert', 'Écrit de haut en bas'],
			]),
		);
		expect(report.pages[0]!.notes).toContain(
			'Rotated text (2 lines) was put in text boxes turned like the original.',
		);
		expect(paragraphs(model.blocks).map(textOf).join(' ')).not.toContain('Écrit');
	});
});

describe('importPdf structure and notes in handmade files', () => {
	it('gives a tagged figure its alternate text and keeps unmatched small text in the body', async () => {
		const bytes = buildPdf([
			'<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 7 0 R >>',
			'<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >>',
			'<< /Type /Page /Parent 2 0 R /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> /XObject << /Im 6 0 R >> >> >>',
			stream(
				'',
				'/Figure << /MCID 0 >> BDC q 200 0 0 100 72 500 cm /Im Do Q EMC ' +
					'BT /F1 12 Tf 72 700 Td (Body text without a mark.) Tj ET ' +
					'BT /F1 8 Tf 72 60 Td (1 Small print that is not a note.) Tj ET',
			),
			'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
			stream(
				'/Type /XObject /Subtype /Image /Width 2 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8',
				'\u0000\u00ff',
				true,
			),
			'<< /Type /StructTreeRoot /K << /S /Figure /Alt (Sales by region) /Pg 3 0 R /K 0 >> >>',
		]);
		const { model } = await importPdf(bytes);
		const picture = paragraphs(model.blocks)
			.flatMap((paragraph) => paragraph.runs)
			.find((run) => run.image)?.image;
		expect(picture?.altText).toBe('Sales by region');
		expect(paragraphs(model.blocks).map(textOf)).toContain('1 Small print that is not a note.');
		expect(model.footnotes ?? []).toHaveLength(0);
	});

	it('takes a footnote with a symbol mark and keeps the symbol', async () => {
		const { model } = await importPdf(
			simplePdf(
				'BT /F1 12 Tf 72 700 Td (A claim) Tj 6 Ts /F1 7 Tf (*) Tj 0 Ts /F1 12 Tf ( and more text, long enough to be the body of the page.) Tj ET ' +
					'72 80 m 200 80 l S ' +
					'BT /F1 5 Tf 72 66 Td 3 Ts (*) Tj 0 Ts /F1 9 Tf ( The source of the claim.) Tj ET',
			),
		);
		const runs = paragraphs(model.blocks).flatMap((paragraph) => paragraph.runs);
		expect(runs.find((run) => run.noteReference)?.noteReference?.id).toBe('1');
		// The symbol follows the reference as Word's custom mark, in the body and in the note.
		expect(paragraphs(model.blocks).map(textOf)).toEqual([
			'A claim* and more text, long enough to be the body of the page.',
		]);
		expect(paragraphs(model.footnotes![0]!.blocks).map(textOf).join('')).toBe(
			'* The source of the claim.',
		);
	});
});
