import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { layoutDocumentModel, createFakeMeasurer } from './layout/index';
import { loadDocx, type DocumentModel, type InlineImage } from './index';
import { shapeColorCss, shapeDefaultTextColor, shapeSvg } from './shape-svg';
import { shapePaint } from './layout/shape-paint';
import { simpleTextBoxDrawing } from './text-box-drawing';
import { expectParagraph, must } from './test-support/access';
import { drawingsFixture } from './test-support/drawings-fixture';

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
