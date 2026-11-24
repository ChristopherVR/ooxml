import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../parse';
import type { Block, Paragraph } from '../model';
import { simplePdf } from '../../pdf/test-support/build';
import { importPdf, isPdf } from './index';
import { PDF_IMPORT_NOTICE } from './import';

const fixture = async (name: string) =>
	new Uint8Array(await readFile(new URL(`./__fixtures__/${name}`, import.meta.url)));
const paragraphs = (blocks: Block[]) =>
	blocks.filter((block): block is Paragraph => block.type === 'paragraph');
const textOf = (paragraph: Paragraph) => paragraph.runs.map((run) => run.text).join('');
const words = (text: string) => text.replace(/\s+/g, ' ').trim();

describe('importPdf on a Word PDF (rapport.pdf, exported by Word 16)', async () => {
	const source = await loadDocx(await fixture('source/rapport.docx'));
	const { model, report, document } = await importPdf(await fixture('rapport.pdf'));
	const body = paragraphs(model.blocks);

	it('keeps the text in reading order, as flowing paragraphs', () => {
		const expected = words(paragraphs(source.model.blocks).map(textOf).join(' '));
		expect(words(body.map(textOf).join(' '))).toBe(expected);
		// The long paragraph that crosses the page break is one paragraph again.
		expect(body.map(textOf).filter((text) => text.startsWith('4. '))).toHaveLength(1);
		expect(body.find((paragraph) => textOf(paragraph).startsWith('4. '))!.align).toBe('justify');
	});

	it('keeps character formatting', () => {
		const runs = body.find((paragraph) => textOf(paragraph).startsWith('Un passage'))!.runs;
		expect(runs.find((run) => run.text === 'gras')).toMatchObject({ bold: true });
		expect(runs.find((run) => run.text === 'italique')).toMatchObject({ italic: true });
		expect(runs.find((run) => run.text === 'souligné')?.underline).toBeTruthy();
		expect(runs.find((run) => run.text === 'rouge')?.color?.toUpperCase()).toBe('#C00000');
		expect(runs.find((run) => run.text === '2')?.verticalAlign).toBe('superscript');
	});

	it('makes headings, the centred caption and the picture', () => {
		expect(body.filter((paragraph) => paragraph.style?.startsWith('Heading')).map(textOf)).toEqual([
			'Rapport annuel',
			'Activité',
			'Perspectives',
		]);
		expect(body.find((paragraph) => textOf(paragraph).startsWith('Texte centré'))!.align).toBe(
			'center',
		);
		const pictures = body.flatMap((paragraph) => paragraph.runs.filter((run) => run.image));
		expect(pictures).toHaveLength(1);
		expect(document.media?.size).toBe(1);
	});

	it('moves the running head and foot to the header and footer, with page fields', () => {
		expect(
			body
				.map(textOf)
				.some((text) => text.includes('Rapport annuel 2026') || /^Page \d/.test(text)),
		).toBe(false);
		const section = model.sections![0]!;
		const header = paragraphs(section.headers!.default!.blocks);
		const footer = paragraphs(section.footers!.default!.blocks);
		expect(header.map(textOf)).toEqual(['Société Exemple – Rapport annuel 2026']);
		expect(header[0]!.align).toBe('right');
		expect(footer[0]!.align).toBe('center');
		const fields = footer[0]!.runs.filter((run) => run.field).map((run) => run.field!.instr.trim());
		expect(fields).toEqual(['PAGE', 'NUMPAGES']);
	});

	it('keeps the page size and margins and the document properties', () => {
		const section = model.sections![0]!;
		expect(section.pageWidthTwips).toBe(11906);
		expect(section.pageHeightTwips).toBe(16838);
		expect(Math.abs(section.marginLeftTwips - 1417)).toBeLessThan(40);
		expect(model.properties).toMatchObject({ title: 'Rapport annuel', creator: 'Exemple' });
	});

	it('says it was converted and saves a .docx that opens again', async () => {
		expect(model.warnings[0]).toBe(PDF_IMPORT_NOTICE);
		expect(report).toMatchObject({ encrypted: false, signed: false, fonts: ['Calibri'] });
		const reopened = await loadDocx(await document.save());
		expect(words(paragraphs(reopened.model.blocks).map(textOf).join(' '))).toBe(
			words(body.map(textOf).join(' ')),
		);
	});
});

describe('importPdf', () => {
	it('recognises PDF bytes', async () => {
		expect(isPdf(await fixture('lettre.pdf'))).toBe(true);
		expect(isPdf(new TextEncoder().encode('PK\u0003\u0004'))).toBe(false);
	});

	it('opens encrypted files and reports their permissions', async () => {
		const { model, report } = await importPdf(await fixture('lettre-aes256.pdf'), {
			password: 'secret',
		});
		expect(paragraphs(model.blocks).map(textOf)).toContain('Corps de la lettre.');
		expect(report.encrypted).toBe(true);
		expect(report.permissions).toMatchObject({ copy: true, modify: false });
		await expect(importPdf(await fixture('lettre-aes256.pdf'))).rejects.toMatchObject({
			code: 'password-required',
		});
	});

	it('honours a page range and reports rotated text', async () => {
		const { model, report } = await importPdf(await fixture('rapport.pdf'), {
			pages: { from: 2, to: 2 },
		});
		// One page has no repeated head or foot: they stay in the text.
		expect(paragraphs(model.blocks).map(textOf).slice(0, 2)).toEqual([
			'Société Exemple – Rapport annuel 2026',
			expect.stringMatching(/^assez long/),
		]);
		expect(report.pages.map((page) => page.index)).toEqual([1]);
		const rotated = await importPdf(
			simplePdf('BT /F1 12 Tf 72 700 Td (Flat) Tj 0 1 -1 0 300 300 Tm (Turned) Tj ET'),
		);
		expect(paragraphs(rotated.model.blocks).map(textOf)).toEqual(['Flat']);
		const box = paragraphs(rotated.model.blocks)[0]!.runs.find((run) => run.image)!.image!;
		expect(box.textBoxContent?.vertical).toBe('vert270');
		expect(paragraphs(box.textBoxContent!.blocks).map(textOf)).toEqual(['Turned']);
		expect(rotated.report.pages[0]!.notes).toContain(
			'Rotated text (1 line) was put in text boxes turned like the original.',
		);
	});

	it('keeps a scanned page as a picture unless OCR is supplied', async () => {
		const image =
			'<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Length 4 >>\nstream\n\u0000\u00ff\u00ff\u0000\nendstream';
		const scan = simplePdf('q 612 0 0 792 0 0 cm /Im Do Q', {
			resources: '/XObject << /Im 6 0 R >>',
			objects: [image],
		});
		const kept = await importPdf(scan);
		expect(kept.report.pages[0]).toMatchObject({ scanned: true });
		expect(
			paragraphs(kept.model.blocks)
				.flatMap((paragraph) => paragraph.runs)
				.some((run) => run.image),
		).toBe(true);
		const read = await importPdf(scan, {
			ocr: async ({ width }) => [
				{ text: 'Recognised', x: 72, y: 100, width: width / 4, height: 12 },
			],
		});
		expect(paragraphs(read.model.blocks).map(textOf)).toEqual(['Recognised']);
	});
});
