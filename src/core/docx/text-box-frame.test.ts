import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { layoutDocumentModel, createFakeMeasurer } from './layout/index';
import { loadDocx, type DocumentModel, type InlineImage } from './index';
import { shapeColorCss, shapeDefaultTextColor, shapeSvg } from './shape-svg';
import { shapePaint } from './layout/shape-paint';
import { simpleTextBoxDrawing } from './text-box-drawing';
import { expectParagraph, must } from './test-support/access';
import { drawingsFixture } from './test-support/drawings-fixture';
import { parseShape } from './shape-parse';
import { parseXml } from '../xml/xml';

const bodyDrawings = (model: DocumentModel): InlineImage[] =>
	model.blocks.flatMap((block) =>
		block.type === 'paragraph' ? block.runs.flatMap((run) => (run.image ? [run.image] : [])) : [],
	);
const byName = (model: DocumentModel, text: string): InlineImage =>
	must(
		bodyDrawings(model).find((image) =>
			image.textBoxContent?.blocks.some((block) =>
				expectParagraph(block).runs.some((run) => run.text === text),
			),
		),
		text,
	);
const documentXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;

describe('text boxes written with txBox="1"', () => {
	it('are read with their frame, so they are drawn like any shape, and stay text boxes', async () => {
		const { model } = await loadDocx(await drawingsFixture());
		const floating = byName(model, 'Floating note');
		expect(floating).toMatchObject({
			unsupported: 'Text box',
			anchored: true,
			textBoxText: ['Floating note', 'Anchored text box'],
			shape: {
				preset: 'rect',
				fill: { kind: 'solid', color: { value: 'FFF2CC' } },
				line: { widthEmu: 12700 },
				editable: false,
				textEditable: false,
			},
		});
		expect(floating.shape).not.toHaveProperty('text');
		const inline = byName(model, 'Inline box');
		expect(inline).toMatchObject({ textBoxEditable: true, textBoxBorder: true });
		expect(inline.shape?.editable).toBe(false);
	});

	it('are drawn as floats and inline objects in Print Layout with their paragraphs', async () => {
		const { model } = await loadDocx(await drawingsFixture());
		const result = layoutDocumentModel(model, createFakeMeasurer());
		const floats = result.pages.flatMap((page) => page.floats ?? []);
		expect(floats).toHaveLength(1);
		expect(floats[0]?.chart?.shape).toBe(true);
		expect(floats[0]?.textBox?.blocks).toHaveLength(2);
		const objects = result.pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'paragraph'
						? block.lines.flatMap((line) => line.fragments.flatMap((f) => f.object ?? []))
						: [],
				),
			),
		);
		expect(objects.filter((object) => object.textBox)).toHaveLength(2);
		expect(result.approximations.join(' ')).not.toMatch(/Text boxes/);
	});

	it('save byte-identically when unchanged', async () => {
		const bytes = await drawingsFixture();
		const loaded = await loadDocx(bytes);
		const saved = await loaded.save(structuredClone(loaded.model));
		expect(await documentXml(saved)).toBe(await documentXml(bytes));
		const header = async (data: Uint8Array) =>
			(await JSZip.loadAsync(data)).file('word/header1.xml')!.async('string');
		expect(await header(saved)).toBe(await header(bytes));
	});

	it('are drawn after an edit as the edit is written', () => {
		const drawn = simpleTextBoxDrawing(['One', 'Two'], false);
		expect(drawn.shape).toMatchObject({
			preset: 'rect',
			fill: { kind: 'solid' },
			line: { fill: { kind: 'none' } },
			editable: false,
		});
		expect(
			drawn.textBoxContent.blocks.map((block) => expectParagraph(block).runs[0]?.text),
		).toEqual(['One', 'Two']);
	});
});

describe('shape text colour', () => {
	it('comes from the style font reference, in SVG text and over text box paragraphs alike', async () => {
		const { model } = await loadDocx(await drawingsFixture());
		const styled = must(
			bodyDrawings(model).find((image) => image.shape?.fontColor),
			'styled shape',
		);
		const shape = must(styled.shape, 'shape');
		const grey = must(shapeColorCss(shape.fontColor!, model.theme), 'font colour');
		expect(grey).toMatch(/^#[0-9a-f]{6}$/i);
		expect(grey.toUpperCase()).not.toBe('#000000');
		expect(grey.toUpperCase()).not.toBe('#FFFFFF');
		expect(shapeDefaultTextColor(shape, model.theme)).toBe(grey);
		expect(shapePaint(styled, model.theme).chart?.textColor).toBe(grey);
		expect(shapeSvg(shape, 100, 40, model.theme)).toContain(`fill="${grey}"`);
	});

	it('falls back to the fill brightness without a font reference', async () => {
		const { model } = await loadDocx(await drawingsFixture());
		const floating = byName(model, 'Floating note');
		expect(floating.shape?.fontColor).toBeUndefined();
		expect(shapePaint(floating, model.theme).chart?.textColor).toBe('#000000');
	});
});

describe('shape text colour on save', () => {
	const styledText = (xml: string, text = 'Styled text') => {
		const shapes = [...xml.matchAll(/<wps:wsp>.*?<\/wps:wsp>/gs)].map((match) => match[0]);
		return must(
			shapes.find((shape) => shape.includes(text)),
			'styled shape XML',
		);
	};

	it('writes no run colour where the style font reference colours the text', async () => {
		const loaded = await loadDocx(await drawingsFixture());
		const model = structuredClone(loaded.model);
		const styled = must(
			bodyDrawings(model).find((image) => image.shape?.fontColor),
			'styled shape',
		);
		expect(styled.shape).toMatchObject({ editable: true, textEditable: true });
		expect(styled.shape?.textColor).toBeUndefined();
		// Resize and refill: the shape is rebuilt and its text rewritten.
		styled.widthPx += 20;
		styled.shape = {
			...styled.shape!,
			fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFFFCC', transforms: [] } },
		};
		const saved = await loaded.save(model);
		const xml = styledText(await documentXml(saved));
		expect(xml).toContain('Styled text');
		expect(xml).toContain('FFFFCC');
		expect(xml).not.toContain('<w:color');
		const reopened = await loadDocx(saved);
		const again = must(
			bodyDrawings(reopened.model).find((image) => image.shape?.fontColor),
			'reopened styled shape',
		);
		const grey = shapeColorCss(again.shape!.fontColor!, reopened.model.theme);
		expect(shapeDefaultTextColor(again.shape!, reopened.model.theme)).toBe(grey);
		expect(shapePaint(again, reopened.model.theme).chart?.textColor).toBe(grey);
	});

	it("keeps the runs' own colour when the text is rewritten, and leaves mixed colours alone", async () => {
		const red = (await documentXml(await drawingsFixture())).replace(
			'<w:r><w:t>Styled text</w:t></w:r>',
			'<w:r><w:rPr><w:color w:val="C00000"/></w:rPr><w:t>Styled text</w:t></w:r>',
		);
		const zip = await JSZip.loadAsync(await drawingsFixture());
		zip.file('word/document.xml', red);
		const coloured = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const model = structuredClone(coloured.model);
		const styled = must(
			bodyDrawings(model).find((image) => image.shape?.fontColor),
			'styled shape',
		);
		expect(styled.shape).toMatchObject({ textColor: 'C00000', textEditable: true });
		expect(shapeSvg({ ...styled.shape!, text: ['Styled text'] }, 100, 40)).toContain(
			'fill="#C00000"',
		);
		styled.shape = { ...styled.shape!, text: ['Edited'] };
		const xml = styledText(await documentXml(await coloured.save(model)), 'Edited');
		expect(xml).toContain('<w:color w:val="C00000"/>');
		expect(xml).toContain('Edited');
		const mixed = red.replace(
			'<w:t>Styled text</w:t></w:r>',
			'<w:t>Styled</w:t></w:r><w:r><w:t> text</w:t></w:r>',
		);
		zip.file('word/document.xml', mixed);
		const reread = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(
			bodyDrawings(reread.model).find((image) => image.shape?.fontColor)?.shape?.textEditable,
		).toBe(false);
	});
});

describe('outlines and fills a shape does not set itself', () => {
	const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
	const shape = (spPr: string, style = '') =>
		must(
			parseShape(
				parseXml(
					`<a:graphicData xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w="${W}" uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPr}</wps:spPr>${style}<wps:bodyPr/></wps:wsp></a:graphicData>`,
				).documentElement,
			),
			'shape',
		);
	const strokes = (svg: string | undefined) =>
		[...(svg ?? '').matchAll(/stroke="([^"]*)"/g)].map((m) => m[1]);
	const fills = (svg: string | undefined) =>
		[...(svg ?? '').matchAll(/ fill="([^"]*)"/g)].map((m) => m[1]);
	const style = (ln: string, fill = '1') =>
		`<wps:style><a:lnRef idx="${ln}"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="${fill}"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>`;

	it('draws no outline for an unstyled shape whose outline sets only a width', () => {
		const spec = shape('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="12700"/>');
		expect(spec.line).toMatchObject({ widthEmu: 12700, fill: { kind: 'none' } });
		expect(strokes(shapeSvg(spec, 40, 20))).toEqual(['none']);
	});

	it('draws no outline or fill for a style reference with idx="0"', () => {
		const spec = shape('', style('0', '0'));
		expect(strokes(shapeSvg(spec, 40, 20))).toEqual(['none']);
		expect(fills(shapeSvg(spec, 40, 20))).toEqual(['none']);
	});

	it("draws the style's outline and fill for a non-zero reference", () => {
		const spec = shape('', style('2'));
		expect(strokes(shapeSvg(spec, 40, 20))).not.toEqual(['none']);
		expect(fills(shapeSvg(spec, 40, 20))).not.toEqual(['none']);
	});
});
