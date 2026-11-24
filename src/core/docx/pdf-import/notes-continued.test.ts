import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../parse';
import type { Block, Paragraph } from '../model';
import { buildPdf, stream } from '../../pdf/test-support/build';
import { convertPdf, importPdf } from './index';

const fixture = async (name: string) =>
	new Uint8Array(await readFile(new URL(`./__fixtures__/${name}`, import.meta.url)));
const paragraphs = (blocks: Block[]) =>
	blocks.filter((block): block is Paragraph => block.type === 'paragraph');
const textOf = (paragraph: Paragraph) => paragraph.runs.map((run) => run.text).join('');
const noteText = (blocks: Block[]) => paragraphs(blocks).map(textOf);

describe('importPdf on a footnote Word continues on the next page (notes-suite.pdf, Word 16)', async () => {
	const { model, document } = await importPdf(await fixture('notes-suite.pdf'));
	const body = paragraphs(model.blocks);
	const notes = model.footnotes ?? [];

	it('joins the continuation on page 2 to the note it goes on with', () => {
		expect(notes).toHaveLength(2);
		const [first, second] = notes.map((note) => noteText(note.blocks));
		// The note's own two paragraphs; the second runs over the page break as one.
		expect(first).toHaveLength(2);
		expect(first![0]!.trim()).toMatch(/^Cette note est volontairement longue\./);
		expect(first![1]).toMatch(/^Son second paragraphe/);
		expect(first![1]).toMatch(/sous le séparateur de continuation\.$/);
		expect(first![1]).not.toMatch(/ {2}/);
		expect(second!.join('').trim()).toBe('La note courte de la deuxième page.');
	});

	it('keeps the continuation out of the body and references both notes', () => {
		const text = body.map(textOf).join(' ');
		expect(text).not.toContain('séparateur de continuation');
		expect(text).not.toContain('La note courte');
		const references = body.flatMap((paragraph) =>
			paragraph.runs.filter((run) => run.noteReference).map((run) => run.noteReference!.id),
		);
		expect(references).toEqual(['1', '2']);
	});

	it('saves both notes', async () => {
		const reopened = await loadDocx(await document.save());
		expect(reopened.model.footnotes).toHaveLength(2);
	});
});

describe('importPdf on endnotes collected at the end (notes-fin.pdf, Word 16)', async () => {
	const result = await convertPdf(await fixture('notes-fin.pdf'));
	const { model } = await loadDocx(result.docx);
	const body = paragraphs(model.blocks);

	it('makes the notes after the separator real endnotes in Word’s i, ii, iii numbering', () => {
		expect((model.endnotes ?? []).map((note) => noteText(note.blocks).join('').trim())).toEqual([
			'Première note de fin, rassemblée à la fin du document.',
			'Deuxième note de fin, assez longue pour occuper deux lignes de la zone des notes, avec une explication de détail.',
			'Troisième note de fin.',
		]);
		expect(model.footnotes ?? []).toHaveLength(0);
		expect(model.endnoteNumFmt).toBe('lowerRoman');
	});

	it('turns the raised marks on both pages into endnote references', () => {
		const references = body.flatMap((paragraph) =>
			paragraph.runs
				.filter((run) => run.noteReference)
				.map((run) => [run.noteReference!.kind, run.noteReference!.id]),
		);
		expect(references).toEqual([
			['endnote', '1'],
			['endnote', '2'],
			['endnote', '3'],
		]);
		const text = body.map(textOf).join(' ');
		expect(text).not.toContain('Première note de fin');
		expect(text).toContain('La conclusion du document.');
	});

	it('writes word/endnotes.xml with the separators, its relationship, type and styles', async () => {
		const zip = await JSZip.loadAsync(result.docx);
		const endnotes = await zip.file('word/endnotes.xml')!.async('string');
		expect(endnotes).toContain('<w:endnote w:type="separator" w:id="-1">');
		expect(endnotes).toContain('<w:endnote w:type="continuationSeparator" w:id="0">');
		expect(endnotes).toContain('<w:endnoteRef/>');
		expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toMatch(
			/relationships\/endnotes" Target="endnotes\.xml"/,
		);
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"',
		);
		expect(await zip.file('word/settings.xml')!.async('string')).toContain(
			'<w:endnotePr><w:numFmt w:val="lowerRoman"/><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>',
		);
		const styles = await zip.file('word/styles.xml')!.async('string');
		expect(styles).toContain('w:styleId="EndnoteText"');
		expect(styles).toContain('w:styleId="EndnoteReference"');
	});
});

/** A PDF of `pages` content streams, Helvetica as /F1, US Letter. */
function pagesPdf(contents: string[]): Uint8Array {
	const count = contents.length;
	const kids = contents.map((_, index) => `${3 + index * 2} 0 R`).join(' ');
	const font = 3 + count * 2;
	return buildPdf([
		'<< /Type /Catalog /Pages 2 0 R >>',
		`<< /Type /Pages /Kids [${kids}] /Count ${count} /MediaBox [0 0 612 792] >>`,
		...contents.flatMap((content, index) => [
			`<< /Type /Page /Parent 2 0 R /Contents ${4 + index * 2} 0 R /Resources << /Font << /F1 ${font} 0 R >> >> >>`,
			stream('', content),
		]),
		'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
	]);
}

const line = (y: number, text: string, size = 12) =>
	`BT /F1 ${size} Tf 72 ${y} Td (${text}) Tj ET `;
const raised = (y: number, before: string, mark: string, after: string) =>
	`BT /F1 12 Tf 72 ${y} Td (${before}) Tj 6 Ts /F1 7 Tf (${mark}) Tj 0 Ts /F1 12 Tf (${after}) Tj ET `;
const noteLine = (y: number, mark: string, text: string) =>
	`BT /F1 5 Tf 72 ${y} Td 3 Ts (${mark}) Tj 0 Ts /F1 9 Tf ( ${text}) Tj ET `;
const filler = (from: number, count: number) =>
	Array.from({ length: count }, (_, index) =>
		line(from - index * 16, `Body text line ${index + 1} that keeps the page going on and on.`),
	).join('');

describe('importPdf on notes across pages in handmade files', () => {
	it('starts a note on the next page when its mark sits at the foot of a full page', async () => {
		const { model } = await importPdf(
			pagesPdf([
				filler(720, 20) + raised(390, 'The last line ends with a claim', '1', ' at the foot.'),
				filler(720, 4) +
					'72 92 m 216 92 l S ' +
					noteLine(80, '1', 'The note of the claim on the page before.') +
					noteLine(68, '2', 'A note of this page.') +
					raised(640, 'This page has its own claim', '2', ' too.'),
			]),
		);
		expect((model.footnotes ?? []).map((note) => noteText(note.blocks).join('').trim())).toEqual([
			'The note of the claim on the page before.',
			'A note of this page.',
		]);
		const references = paragraphs(model.blocks).flatMap((paragraph) =>
			paragraph.runs.filter((run) => run.noteReference).map((run) => run.noteReference!.id),
		);
		expect(references).toEqual(['1', '2']);
	});

	it('takes numbered notes after a Notes heading as endnotes', async () => {
		const { model } = await importPdf(
			pagesPdf([
				filler(720, 6) + raised(600, 'A first claim', '1', ' in the text.'),
				raised(720, 'A second claim', '2', ' later on.') +
					filler(700, 3) +
					line(600, 'Notes', 14) +
					line(580, '1. The source of the first claim.', 10) +
					line(566, '2. The source of the second claim, long enough to fill its line,', 10) +
					line(552, 'which takes two lines.', 10),
			]),
		);
		expect((model.endnotes ?? []).map((note) => noteText(note.blocks).join('').trim())).toEqual([
			'The source of the first claim.',
			'The source of the second claim, long enough to fill its line, which takes two lines.',
		]);
		expect(model.endnoteNumFmt).toBe('decimal');
		expect(paragraphs(model.blocks).map(textOf)).not.toContain('Notes');
	});

	it('leaves footnotes at the foot of the last page as footnotes', async () => {
		const { model } = await importPdf(
			pagesPdf([
				filler(720, 3) +
					raised(660, 'A claim', '1', ' on the only page.') +
					filler(640, 30) +
					'72 92 m 216 92 l S ' +
					noteLine(80, '1', 'Its footnote.'),
			]),
		);
		expect(model.endnotes ?? []).toHaveLength(0);
		expect(model.footnotes ?? []).toHaveLength(1);
	});
});
