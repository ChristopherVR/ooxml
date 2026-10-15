import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx, saveDocx } from './index';
import type { DocumentModel, InlineImage, Paragraph } from './model';
import { expectParagraph } from './test-support/access';
import { newShape } from './shape-model';
import { shapeSvg } from './shape-svg';
import { createDocument } from './model';

const fixture = new URL('./__fixtures__/shapes/shapes.docx', import.meta.url);
const load = async () => loadDocx(new Uint8Array(await readFile(fixture)));
const xmlOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;
const shapes = (paragraph: Paragraph): InlineImage[] =>
	paragraph.runs.flatMap((run) => (run.image?.shape ? [run.image] : []));

describe('shapes authored by Word', () => {
	it('reads preset geometry, fill, outline, text, size and floating position', async () => {
		const { model } = await load();
		const found = shapes(expectParagraph(model.blocks[0]));
		const byPreset = (preset: string) => found.find((image) => image.shape?.preset === preset)!;
		expect(found.map((image) => image.shape?.preset).sort()).toEqual([
			'ellipse',
			'line',
			'rect',
			'rightArrow',
		]);
		const rect = byPreset('rect');
		expect(rect.shape).toMatchObject({
			fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
			line: { widthEmu: 38100 },
			text: ['Red box'],
			textEditable: true,
		});
		expect(rect).toMatchObject({ unsupported: 'Shape', anchored: true, widthPx: 187 });
		expect(rect.placement).toMatchObject({ wrap: 'none', relativeFromV: 'paragraph' });
		expect(byPreset('rightArrow').placement?.wrap).toBe('square');
		expect(byPreset('ellipse').shape?.fill).toEqual({ kind: 'none' });
	});

	it('keeps the drawings byte-for-byte when the paragraph text is edited', async () => {
		const loaded = await load();
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs[paragraph.runs.length - 1]!.text = 'Edited';
		const xml = await xmlOf(await loaded.save(loaded.model));
		expect(xml).toContain('<w:t>Edited</w:t>');
		expect(xml.match(/<mc:AlternateContent>/g)).toHaveLength(4);
		expect(xml).toContain('prst="rightArrow"');
		expect(xml).toContain('<v:rect');
	});

	it('rebuilds an edited shape as a modern drawing and keeps its identity and style', async () => {
		const loaded = await load();
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const rect = shapes(paragraph).find((image) => image.shape?.preset === 'rect')!;
		const before = await xmlOf(await loaded.save(loaded.model));
		const docPr = /<wp:docPr id="(\d+)" name="Rectangle 1"/.exec(before)!;
		rect.shape = {
			...rect.shape!,
			preset: 'roundRect',
			fill: { kind: 'solid', color: { kind: 'srgb', value: '112233', transforms: [] } },
			text: ['Changed text'],
		};
		rect.widthPx = 200;
		rect.placement = { ...rect.placement!, wrap: 'topAndBottom', offsetXPx: 10 };
		const xml = await xmlOf(await loaded.save(loaded.model));
		expect(xml).toContain('prst="roundRect"');
		expect(xml).toContain('val="112233"');
		expect(xml).toContain('<w:t>Changed text</w:t>');
		expect(xml).toContain('<wp:wrapTopAndBottom/>');
		expect(xml).toContain(`<wp:docPr id="${docPr[1]}" name="Rectangle 1"`);
		expect(xml).toContain('<wp:posOffset>95250</wp:posOffset>');
		expect(xml).not.toContain('<v:rect');
		expect(xml.match(/<mc:AlternateContent>/g)).toHaveLength(3);
		const reopened = await loadDocx(await loaded.save(loaded.model));
		const again = shapes(expectParagraph(reopened.model.blocks[0])).find(
			(image) => image.shape?.preset === 'roundRect',
		)!;
		expect(again).toMatchObject({
			widthPx: 200,
			placement: { wrap: 'topAndBottom', offsetXPx: 10 },
		});
	});
});

describe('new shapes', () => {
	const withShape = (image: InlineImage): DocumentModel => {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Text ' }, { text: '', image }] }];
		return model;
	};

	it('writes a floating shape with the default style and reads it back', async () => {
		const model = withShape({
			relId: '',
			partName: '',
			contentType: 'application/octet-stream',
			widthPx: 144,
			heightPx: 72,
			anchored: true,
			unsupported: 'Shape',
			placement: {
				wrap: 'none',
				relativeFrom: 'column',
				offsetXPx: 30,
				relativeFromV: 'paragraph',
				offsetYPx: 12,
			},
			shape: { ...newShape('rightArrow'), text: ['Go'] },
		});
		const bytes = await saveDocx(model);
		const xml = await xmlOf(bytes);
		expect(xml).toContain('<wp:anchor');
		expect(xml).toContain(
			'<wp:positionH relativeFrom="column"><wp:posOffset>285750</wp:posOffset>',
		);
		expect(xml).toContain(
			'<wp:positionV relativeFrom="paragraph"><wp:posOffset>114300</wp:posOffset>',
		);
		expect(xml).toContain('<wp:wrapNone/>');
		expect(xml).toContain('<a:prstGeom prst="rightArrow">');
		expect(xml).toContain('<wps:style>');
		expect(xml).toContain('<w:t>Go</w:t>');
		expect(xml).not.toContain('<a:solidFill><a:srgbClr');
		const reopened = await loadDocx(bytes);
		const [shape] = shapes(expectParagraph(reopened.model.blocks[0]));
		expect(shape).toMatchObject({
			anchored: true,
			widthPx: 144,
			heightPx: 72,
			shape: { preset: 'rightArrow', text: ['Go'], editable: true },
		});
	});

	it('writes explicit fill, outline and an inline shape', async () => {
		const model = withShape({
			relId: '',
			partName: '',
			contentType: 'application/octet-stream',
			widthPx: 96,
			heightPx: 96,
			unsupported: 'Shape',
			shape: {
				...newShape('ellipse'),
				fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFFF00', transforms: [] } },
				line: { widthEmu: 28575, fill: { kind: 'none' } },
				text: ['Light'],
			},
		});
		const xml = await xmlOf(await saveDocx(model));
		expect(xml).toContain('<wp:inline');
		expect(xml).toContain('val="FFFF00"/>');
		expect(xml).toMatch(/<a:ln w="28575"><a:noFill\/><\/a:ln>/);
		// Text on a light explicit fill is black so it stays readable.
		expect(xml).toContain('<w:color w:val="000000"/>');
	});
});

describe('shape drawing', () => {
	it('draws preset geometry with the default style, an explicit fill and text', () => {
		const plain = shapeSvg(newShape('rect'), 100, 50)!;
		expect(plain).toContain('fill="#4472C4"');
		expect(plain).toContain('stroke="#2F528F"');
		expect(plain).toContain('viewBox="0 0 100 50"');
		const custom = shapeSvg(
			{
				...newShape('roundRect'),
				fill: { kind: 'solid', color: { kind: 'srgb', value: '00FF00', transforms: [] } },
				line: { widthEmu: 19050, fill: { kind: 'none' } },
				text: ['Hello world'],
			},
			120,
			60,
		)!;
		expect(custom).toContain('fill="#00FF00"');
		expect(custom).toContain('stroke="none"');
		expect(custom).toContain('>Hello world<');
		expect(custom).toContain('fill="#000000"');
		expect(shapeSvg(newShape('notAShape'), 10, 10)).toBeUndefined();
	});

	it('draws lines unfilled and every gallery family', () => {
		const line = shapeSvg(newShape('line'), 100, 40)!;
		expect(line).toContain('fill="none"');
		for (const preset of [
			'rect',
			'ellipse',
			'triangle',
			'rightArrow',
			'flowChartProcess',
			'flowChartDecision',
			'wedgeRectCallout',
			'star5',
			'straightConnector1',
		])
			expect(shapeSvg(newShape(preset), 80, 60), preset).toContain('<path');
	});
});
