import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type DocumentModel, type InlineImage, type Paragraph } from './index';
import { drawingShapeSvg } from './shape-svg';
import { expectParagraph, must } from './test-support/access';
import { letterheadFixture } from './test-support/letterhead-fixture';

const headerParagraph = (model: DocumentModel): Paragraph =>
	expectParagraph(must(model.sections?.[0]?.headers?.first, 'first-page header').blocks[0]);
const drawings = (paragraph: Paragraph): InlineImage[] =>
	paragraph.runs.flatMap((run) => (run.image ? [run.image] : []));
const headerXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/header1.xml')!.async('string')) as string;

describe('text box content', () => {
	it('keeps a rect text box without txBox="1" as real paragraphs with their formatting', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const [box] = drawings(headerParagraph(model));
		expect(box?.shape?.preset).toBe('rect');
		const content = must(box?.textBoxContent, 'text box content');
		expect(content).toMatchObject({
			anchor: 't',
			insetLeftEmu: 91440,
			insetTopEmu: 45720,
		});
		const [name, detail, rule, address] = content.blocks.map(expectParagraph);
		expect(name?.runs[0]).toMatchObject({ text: 'First Partner', bold: true, color: '#1F3864' });
		expect(detail?.runs[0]).toMatchObject({ text: 'Specialist in contract law', fontSize: 9 });
		expect(rule?.borders?.bottom).toMatchObject({ style: 'single' });
		expect(address?.align).toBe('left');
		expect(address?.runs[0]).toMatchObject({ italic: true, fontSize: 9 });
	});

	it('does not offer to rewrite formatted text as plain lines, and draws the shape without it', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const [box] = drawings(headerParagraph(model));
		expect(box?.shape?.textEditable).toBe(false);
		const svg = must(drawingShapeSvg(box!), 'shape SVG');
		expect(svg).not.toContain('<text');
		expect(svg).not.toContain('First Partner');
	});

	it('reads the connector line with its outline colour, width and dash', async () => {
		const { model } = await loadDocx(await letterheadFixture());
		const [, line] = drawings(headerParagraph(model));
		expect(line?.shape).toMatchObject({ preset: 'straightConnector1', line: { widthEmu: 19050 } });
		expect(line?.textBoxContent).toBeUndefined();
		const svg = must(drawingShapeSvg(line!), 'line SVG');
		expect(svg).toMatch(/stroke="#2E75B6"/i);
		expect(svg).toContain('stroke-dasharray');
	});

	it('writes the header part back unchanged when the document is saved', async () => {
		const bytes = await letterheadFixture();
		const source = await headerXml(bytes);
		const loaded = await loadDocx(bytes);
		expect(await headerXml(await loaded.save(structuredClone(loaded.model)))).toBe(source);
	});

	it('keeps the text box XML verbatim when other text in its header paragraph is edited', async () => {
		const bytes = await letterheadFixture();
		const source = await headerXml(bytes);
		const loaded = await loadDocx(bytes);
		const model = structuredClone(loaded.model);
		const paragraph = headerParagraph(model);
		paragraph.runs = paragraph.runs.map((run) =>
			run.text === 'Letterhead' ? { ...run, text: 'Letterhead 2026' } : run,
		);
		const saved = await headerXml(await loaded.save(model));
		expect(saved).toContain('Letterhead 2026');
		const box = /<w:drawing>.*?<\/w:drawing>/s;
		expect(saved.match(box)?.[0]).toBe(source.match(box)?.[0]);
		expect(saved).toContain(
			'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="1F3864"/></w:pBdr>',
		);
	});
});
