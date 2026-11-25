import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { simplePdf } from '../../pdf/test-support/build';
import type { Block, Paragraph } from '../model';
import { schemaErrors } from '../test-support/schema-validation';
import { convertPdf } from './import';
import { importPdf } from './index';

const text = (x: number, y: number, words: string, size = 12) =>
	`BT /F1 ${size} Tf ${x} ${y} Td (${words}) Tj ET `;
// A bar chart: two axes, three bars and a trend line, all touching, in three looks.
const chart =
	'0 0 0 RG 1 w 100 500 m 100 650 l S 100 500 m 400 500 l S ' +
	'0.2 0.4 0.8 rg 120 500 40 80 re f 180 500 40 120 re f 240 500 40 60 re f ' +
	'1 0 0 RG 2 w 140 580 m 200 620 l 260 560 l S ';
const circle =
	'0 0.5 0 rg 450 630 m 466.57 630 480 616.57 480 600 c 480 583.43 466.57 570 450 570 c ' +
	'433.43 570 420 583.43 420 600 c 420 616.57 433.43 630 450 630 c f ';
const cross = '0 G 1 w 430 280 m 470 320 l S 430 320 m 470 280 l S ';
const table =
	'0 G 0.5 w 72 140 m 372 140 l S 72 120 m 372 120 l S 72 100 m 372 100 l S ' +
	'72 100 m 72 140 l S 222 100 m 222 140 l S 372 100 m 372 140 l S ' +
	text(80, 126, 'Cell one') +
	text(230, 126, 'Cell two') +
	text(80, 106, 'Cell three') +
	text(230, 106, 'Cell four');
const content =
	text(72, 720, 'Sales by quarter', 16) +
	chart +
	circle +
	text(72, 400, 'Body text under the chart.') +
	cross +
	text(72, 200, 'Underlined') +
	'0 g 72 197 60 0.75 re f ' +
	table;

/** document.xml with each drawing turned into the DrawingML locked canvas it is equivalent to. */
function asLockedCanvas(xml: string): string {
	const lc = 'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas';
	const group = '<a:nvGrpSpPr><a:cNvPr id="0" name="g"/><a:cNvGrpSpPr/></a:nvGrpSpPr>';
	return xml
		.replace(/<mc:AlternateContent[^>]*><mc:Choice Requires="wp[sg]">/g, '')
		.replace(/<\/mc:Choice><mc:Fallback\/><\/mc:AlternateContent>/g, '')
		.replace(
			/uri="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordprocessing(Shape|Group)"/g,
			`uri="${lc}"`,
		)
		.replace(
			/<wpg:wgp><wpg:cNvGrpSpPr\/><wpg:grpSpPr>/g,
			`<lc:lockedCanvas xmlns:lc="${lc}">${group}<a:grpSpPr>`,
		)
		.replace(/<\/wpg:grpSpPr>/g, '</a:grpSpPr>')
		.replace(/<\/wpg:wgp>/g, '</lc:lockedCanvas>')
		.replace(
			/<wps:wsp><wps:cNvPr (id="\d+" name="[^"]*")\/><wps:cNvSpPr\/>/g,
			'<a:sp><a:nvSpPr><a:cNvPr $1/><a:cNvSpPr/></a:nvSpPr>',
		)
		.replace(
			/<a:graphicData uri="([^"]+)"><wps:wsp><wps:cNvSpPr\/>/g,
			`<a:graphicData uri="$1"><lc:lockedCanvas xmlns:lc="${lc}">${group}<a:grpSpPr/><a:sp><a:nvSpPr><a:cNvPr id="0" name="s"/><a:cNvSpPr/></a:nvSpPr>`,
		)
		.replace(/<wps:txbx>[\s\S]*?<\/wps:txbx>/g, '')
		.replace(
			/<wps:bodyPr(?:\s[^>]*)?\/><\/wps:wsp><\/a:graphicData>/g,
			'</a:sp></lc:lockedCanvas></a:graphicData>',
		)
		.replace(/<wps:bodyPr(?:\s[^>]*)?\/><\/wps:wsp>/g, '</a:sp>')
		.replace(/<(\/?)wps:spPr>/g, '<$1a:spPr>');
}

const paragraphs = (blocks: Block[]) =>
	blocks.filter((block): block is Paragraph => block.type === 'paragraph');

describe('vector drawings', async () => {
	const bytes = simplePdf(content);
	const { docx, report } = await convertPdf(bytes);
	const xml = await (await JSZip.loadAsync(docx)).file('word/document.xml')!.async('string');
	const anchors = xml.match(/<wp:anchor[\s\S]*?<\/wp:anchor>/g) ?? [];

	it('turns the chart, the circle and the cross into three drawings behind the text', () => {
		expect(anchors).toHaveLength(3);
		expect(anchors.every((anchor) => anchor.includes('behindDoc="1"'))).toBe(true);
		expect(report.pages[0]!.notes).toContain(
			'Vector drawings (9 shapes) became Word drawings placed where the PDF drew them, behind the text; text in them stays ordinary text.',
		);
		expect(report.pages[0]!.notes.some((note) => note.includes('left out'))).toBe(false);
	});

	it('keeps the chart as one group: preset rectangles and lines, a custom path for the trend', () => {
		const group = anchors.find((anchor) => anchor.includes('<wpg:wgp>'))!;
		expect(xml).toContain('<mc:Choice Requires="wpg">');
		expect(group.match(/<wps:wsp>/g)).toHaveLength(6);
		expect(group.match(/prst="rect"/g)).toHaveLength(3);
		expect(group.match(/prst="line"/g)).toHaveLength(2);
		expect(group).toContain('<a:solidFill><a:srgbClr val="3366CC"/></a:solidFill>');
		// The trend line: a red open path of two segments, 2 pt wide, without fill.
		expect(group).toMatch(
			/<a:custGeom>.*<a:path w="\d+" h="\d+"><a:moveTo>.*<\/a:moveTo><a:lnTo>.*<\/a:lnTo><a:lnTo>.*<\/a:lnTo><\/a:path>.*<a:noFill\/><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"\/>/,
		);
		// Placed at the chart's top-left corner on the page.
		expect(group).toContain(
			`<wp:positionH relativeFrom="page"><wp:posOffset>${100 * 12700}</wp:posOffset>`,
		);
		expect(group).toContain(
			`<wp:positionV relativeFrom="page"><wp:posOffset>${(792 - 650) * 12700}</wp:posOffset>`,
		);
	});

	it('draws the circle as an ellipse and the cross as one shape with two paths', () => {
		const ellipse = anchors.find((anchor) => anchor.includes('prst="ellipse"'))!;
		expect(ellipse).toContain(`<wp:extent cx="${60 * 12700}" cy="${60 * 12700}"/>`);
		expect(ellipse).toContain('<a:srgbClr val="008000"/>');
		const crossed = anchors.find((anchor) => (anchor.match(/<a:path /g) ?? []).length === 2)!;
		expect(crossed).toContain('<a:noFill/><a:ln w="12700">');
	});

	it('leaves the table rules and the underline to the table and the text', async () => {
		const { model } = await importPdf(bytes);
		expect(model.blocks.some((block) => block.type === 'table')).toBe(true);
		const underlined = paragraphs(model.blocks)
			.flatMap((paragraph) => paragraph.runs)
			.find((run) => run.text === 'Underlined');
		expect(underlined?.underline).toBeTruthy();
		expect(
			paragraphs(model.blocks).some((paragraph) =>
				paragraph.runs.some((run) => run.text.includes('Body text under the chart.')),
			),
		).toBe(true);
	});

	it('writes valid DrawingML', async () => {
		expect(await schemaErrors(asLockedCanvas(xml))).toEqual([]);
	});

	it('opens a page that holds only drawings, and reports shapes over the limit', async () => {
		// 520 short diagonal strokes, far enough apart to stay separate drawings.
		const many = Array.from({ length: 520 }, (_, index) => {
			const [x, y] = [50 + (index % 26) * 20, 100 + Math.floor(index / 26) * 25];
			return `${x} ${y} m ${x + 6} ${y + 4} l S `;
		}).join('');
		const result = await convertPdf(simplePdf(many));
		const notes = result.report.pages[0]!.notes;
		expect(notes).toContain(
			'Vector drawings (500 shapes) became Word drawings placed where the PDF drew them, behind the text; text in them stays ordinary text.',
		);
		expect(notes).toContain(
			'Vector drawings (20 shapes) were left out: a page converts at most 500 shapes (20000 path segments).',
		);
		const page = await (
			await JSZip.loadAsync(result.docx)
		)
			.file('word/document.xml')!
			.async('string');
		expect(page.match(/<wp:anchor /g)?.length).toBe(500);
		expect(await schemaErrors(asLockedCanvas(page))).toEqual([]);
	});
});

describe('vector drawings in a Word PDF (dessins.pdf, exported by Word 16)', async () => {
	const pdf = new Uint8Array(
		await readFile(new URL('./__fixtures__/dessins.pdf', import.meta.url)),
	);
	const { docx, report } = await convertPdf(pdf);
	const xml = await (await JSZip.loadAsync(docx)).file('word/document.xml')!.async('string');
	const anchors = xml.match(/<wp:anchor[\s\S]*?<\/wp:anchor>/g) ?? [];
	const offset = (anchor: string, axis: 'H' | 'V') =>
		Number(
			anchor.match(
				new RegExp(`<wp:position${axis} relativeFrom="page"><wp:posOffset>(-?\\d+)<`),
			)![1],
		) / 12700;

	it('rebuilds the five drawings where Word placed the shapes', () => {
		expect(anchors).toHaveLength(5);
		const boxes = anchors.map((anchor) => [
			Math.round(offset(anchor, 'H')),
			Math.round(offset(anchor, 'V')),
		]);
		// The rectangle, the ellipse and the line, where Word's object model put them.
		expect(boxes).toEqual(
			expect.arrayContaining([
				[72, 150],
				[230, 150],
				[360, 150],
			]),
		);
		expect(anchors.find((anchor) => anchor.includes('prst="ellipse"'))).toBeDefined();
		// The chart's bars, gridlines and legend keys stay one group.
		expect(
			Math.max(...anchors.map((anchor) => (anchor.match(/<wps:wsp>/g) ?? []).length)),
		).toBeGreaterThan(20);
		expect(report.pages[0]!.notes.some((note) => note.includes('left out'))).toBe(false);
	});

	it('writes valid DrawingML', async () => {
		expect(await schemaErrors(asLockedCanvas(xml))).toEqual([]);
	});
});
