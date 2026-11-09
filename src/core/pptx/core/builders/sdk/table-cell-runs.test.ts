/**
 * A new table cell's rich runs are written as given (one `a:r` per run with
 * its own `a:rPr`, paragraphs at `\n`, links with relationships) and read
 * back as the cell's runs; loaded and edited cells save as before.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { TablePptxElement } from '../../types';
import { createTableElement } from './ElementFactory';
import type { TableInput } from './types';

async function save(input: TableInput) {
	const { handler, createSlide } = await PptxHandler.create();
	const slides = [createSlide('Blank').addTable(input).build(), createSlide('Blank').build()];
	const zip = await JSZip.loadAsync(await handler.save(slides));
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
	const rel = (id: string) =>
		new RegExp(`<Relationship [^>]*Id="${id}"[^>]*>`).exec(rels)?.[0] ?? '';
	const cells = xml.match(/<a:tc\b[\s\S]*?<\/a:tc>/g) ?? [];
	return { xml, rel, cells, bytes: await zip.generateAsync({ type: 'uint8array' }) };
}

const runs = (xml: string) => xml.match(/<a:r>[\s\S]*?<\/a:r>/g) ?? [];

describe('table cell runs', () => {
	it('writes every run with its own properties and an external link', async () => {
		const { cells, rel } = await save({
			rows: [
				{
					cells: [
						{
							runs: [
								{ text: 'A', style: { bold: true, fontFamily: 'Roboto' } },
								{
									text: 'B',
									style: { strikethrough: true, link: { url: 'https://example.com/' } },
								},
							],
						},
					],
				},
			],
		});
		const cell = cells[0];
		expect(cell.match(/<a:p>/g)).toHaveLength(1);
		const [a, b] = runs(cell);
		expect(a).toContain('b="1"');
		expect(a).toContain('<a:latin typeface="Roboto"');
		expect(a).not.toContain('sngStrike');
		expect(b).toContain('strike="sngStrike"');
		const id = /<a:hlinkClick r:id="([^"]+)"/.exec(b)?.[1] ?? '';
		expect(rel(id)).toContain('Target="https://example.com/"');
		expect(rel(id)).toContain('TargetMode="External"');
	});

	it('starts a paragraph at a newline and writes a run slide jump', async () => {
		const { cells, rel } = await save({
			rows: [
				{
					cells: [
						{
							runs: [
								{ text: 'one\ntwo ', style: { italic: true } },
								{ text: 'next', style: { link: { slideIndex: 1 } } },
							],
							style: { alignment: 'center', fontSize: 12 },
						},
					],
				},
			],
		});
		const cell = cells[0];
		expect(cell.match(/<a:p>/g)).toHaveLength(2);
		expect(cell.match(/algn="ctr"/g)).toHaveLength(2);
		// cell style.fontSize is points: 12 pt
		for (const run of runs(cell)) {
			expect(run).toContain('sz="1200"');
		}
		const node = /<a:hlinkClick [^>]*>/.exec(cell)?.[0] ?? '';
		expect(node).toContain('action="ppaction://hlinksldjump"');
		const id = /r:id="([^"]+)"/.exec(node)?.[1] ?? '';
		expect(rel(id)).toContain('Target="slide2.xml"');
	});

	it('leaves plain cells as before', async () => {
		const { cells } = await save({ rows: [{ cells: [{ text: 'x', style: { bold: true } }] }] });
		expect(runs(cells[0])).toHaveLength(1);
		expect(runs(cells[0])[0]).toContain('b="1"');
	});

	it('reads the runs back', async () => {
		const { bytes } = await save({
			rows: [
				{
					cells: [
						{
							runs: [
								{ text: 'A', style: { bold: true } },
								{ text: 'B', style: { italic: true, fontSize: 16 } },
							],
						},
					],
				},
			],
		});
		const data = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
		const table = data.slides[0].elements[0] as {
			tableData?: { rows: Array<{ cells: Array<{ text: string; textRuns?: unknown[] }> }> };
		};
		const cell = table.tableData!.rows[0].cells[0];
		expect(cell.text).toBe('AB');
		expect(cell.textRuns).toEqual([
			expect.objectContaining({ text: 'A', bold: true }),
			expect.objectContaining({ text: 'B', italic: true, fontSize: 12 }),
		]);
	});

	it('lets a run override the cell style and keeps fractional point sizes', async () => {
		const { cells } = await save({
			rows: [
				{
					cells: [
						{
							runs: [{ text: 'A', style: { bold: false } }, { text: 'B' }],
							style: { bold: true, fontSize: 10.5 },
						},
					],
				},
			],
		});
		const [a, b] = runs(cells[0]);
		expect(a).not.toContain('b="1"');
		expect(b).toContain('b="1"');
		expect(a).toContain('sz="1050"');
		expect(b).toContain('sz="1050"');
	});

	it('keeps the cell fill, borders and merge of a cell with runs', async () => {
		const { cells } = await save({
			rows: [
				{
					cells: [
						{
							runs: [{ text: 'A', style: { bold: true } }, { text: 'B' }],
							style: { color: '#FF0000' },
							fill: { type: 'solid', color: '#00B050' },
							borders: { bottom: { color: '#203864', width: 2 } },
							gridSpan: 2,
						},
					],
				},
				{ cells: [{ text: 'C' }, { text: 'D' }] },
			],
		});
		const cell = cells[0];
		expect(cell).toContain('gridSpan="2"');
		const tcPr = /<a:tcPr\b[\s\S]*?<\/a:tcPr>/.exec(cell)?.[0] ?? '';
		expect(tcPr).toContain('<a:srgbClr val="00B050">');
		expect(tcPr).toContain('<a:lnB');
		const [a, b] = runs(cell);
		expect(a).toContain('b="1"');
		expect(b).not.toContain('b="1"');
		expect(a).toContain('val="FF0000"');
		expect(b).toContain('val="FF0000"');
	});

	it('writes plain text for an empty run list', async () => {
		const { cells } = await save({ rows: [{ cells: [{ text: 'plain', runs: [] }] }] });
		expect(runs(cells[0])).toHaveLength(1);
		expect(cells[0]).toContain('<a:t>plain</a:t>');
	});

	it('rejects text that disagrees with the runs', () => {
		expect(() =>
			createTableElement({ rows: [{ cells: [{ text: 'x', runs: [{ text: 'y' }] }] }] }),
		).toThrow(RangeError);
		expect(
			createTableElement({
				rows: [{ cells: [{ text: 'AB', runs: [{ text: 'A' }, { text: 'B' }] }] }],
			}).tableData.rows[0]?.cells[0]?.text,
		).toBe('AB');
	});

	it('writes the flat text once the cell text is changed', async () => {
		const el = createTableElement({
			rows: [{ cells: [{ runs: [{ text: 'A', style: { bold: true } }, { text: 'B' }] }] }],
		});
		const cell = el.tableData.rows[0]?.cells[0];
		if (cell) {
			cell.text = 'edited';
		}
		const { handler, createSlide } = await PptxHandler.create();
		const slide = createSlide('Blank').build();
		slide.elements.push(el);
		const zip = await JSZip.loadAsync(await handler.save([slide]));
		const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		const tc = /<a:tc\b[\s\S]*?<\/a:tc>/u.exec(xml)?.[0] ?? '';
		expect(runs(tc)).toHaveLength(1);
		expect(tc).toContain('<a:t>edited</a:t>');
	});

	it('keeps the runs of a reloaded table when it is saved again', async () => {
		const { bytes } = await save({
			rows: [
				{
					cells: [
						{
							runs: [
								{ text: 'A', style: { bold: true } },
								{ text: 'B', style: { link: { url: 'https://example.com/' } } },
							],
						},
					],
				},
			],
		});
		const handler = new PptxHandler();
		const data = await handler.load(bytes.buffer as ArrayBuffer);
		const loaded = data.slides[0]?.elements[0] as TablePptxElement | undefined;
		expect(loaded?.tableData?.rows[0]?.cells[0]?.textSegments).toBeUndefined();
		const again = await JSZip.loadAsync(await handler.save(data.slides));
		const xml = (await again.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		const [a, b] = runs(/<a:tc\b[\s\S]*?<\/a:tc>/u.exec(xml)?.[0] ?? '');
		expect(a).toContain('b="1"');
		expect(b).toContain('<a:hlinkClick');
	});

	it('leaves out and reports a run slide jump past the last slide', async () => {
		const { handler, createSlide } = await PptxHandler.create();
		const slide = createSlide('Blank').addTable({
			rows: [{ cells: [{ runs: [{ text: 'far', style: { link: { slideIndex: 4 } } }] }] }],
		});
		const zip = await JSZip.loadAsync(await handler.save([slide.build()]));
		const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		expect(xml).not.toContain('r:id=""');
		expect(xml).not.toContain('hlinkClick');
		expect(xml).toContain('<a:t>far</a:t>');
		expect(handler.getCompatibilityWarnings().map((w) => w.code)).toContain(
			'SAVE_SLIDE_JUMP_TARGET_MISSING',
		);
	});
});
