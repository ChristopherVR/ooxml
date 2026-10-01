import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { diagramsIn, loadDocx, resolveDiagramColor, type DocxDiagram } from './index.js';
import { at, expectParagraph } from './test-support/access.js';

// Provenance and the build script: scripts/docx/build-smartart-fixture.mjs (PROVENANCE.md).
const fixture = () =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '__fixtures__/smartart.docx')));
const DIAGRAM_PARTS = ['data', 'layout', 'quickStyle', 'colors', 'drawing'].flatMap((stem) =>
	[1, 2].map((index) => `word/diagrams/${stem}${index}.xml`),
);

async function parts(bytes: Uint8Array, names: string[]): Promise<Record<string, string>> {
	const zip = await JSZip.loadAsync(bytes);
	return Object.fromEntries(
		await Promise.all(names.map(async (name) => [name, await zip.file(name)!.async('string')])),
	);
}

async function withoutDiagramDrawing(bytes: Uint8Array, index: number): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	zip.remove(`word/diagrams/drawing${index}.xml`);
	const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
	zip.file(
		'word/_rels/document.xml.rels',
		rels.replace(new RegExp(`<Relationship [^>]*drawing${index}\\.xml"[^>]*/>`, 'u'), ''),
	);
	return zip.generateAsync({ type: 'uint8array' });
}

const diagramsOf = (model: { blocks: Parameters<typeof diagramsIn>[0] }): DocxDiagram[] =>
	diagramsIn(model.blocks);

describe('SmartArt in a Word document', () => {
	it('models inline and floating diagrams with their parts, extent and alt text', async () => {
		const { model } = await loadDocx(fixture());
		const [inline, floating] = diagramsOf(model) as [DocxDiagram, DocxDiagram];
		expect(inline).toMatchObject({
			placement: 'inline',
			extentEmu: { width: 8255000, height: 5080000 },
			name: 'Diagram 1',
			id: '1',
			relationshipIds: {
				dataRelId: 'rId2',
				layoutRelId: 'rId3',
				styleRelId: 'rId4',
				colorsRelId: 'rId5',
			},
			parts: {
				data: { relId: 'rId2', partName: 'word/diagrams/data1.xml' },
				layout: { relId: 'rId3', partName: 'word/diagrams/layout1.xml' },
				quickStyle: { relId: 'rId4', partName: 'word/diagrams/quickStyle1.xml' },
				colors: { relId: 'rId5', partName: 'word/diagrams/colors1.xml' },
				drawing: { relId: 'rId6', partName: 'word/diagrams/drawing1.xml' },
			},
			layout: { family: 'hierarchy' },
			rendering: 'cached-drawing',
			issues: [],
		});
		expect(inline.layout?.uniqueId).toContain('orgChart1');
		expect(inline.colorsId).toContain('colors/accent1_2');
		expect(inline.quickStyleId).toContain('quickstyle/simple1');
		expect(inline.drawing?.shapes).toHaveLength(9);
		expect(inline.nodes.length).toBeGreaterThan(0);
		expect(floating).toMatchObject({
			placement: 'anchor',
			name: 'Diagram 2',
			parts: { drawing: { relId: 'rId11', partName: 'word/diagrams/drawing2.xml' } },
			rendering: 'cached-drawing',
		});
		expect(floating.drawing?.shapes).toHaveLength(13);
	});

	it('keeps the drawing a placeholder run (unsupported SmartArt) with alt text and anchoring', async () => {
		const { model } = await loadDocx(fixture());
		const inline = at(expectParagraph(model.blocks[1]).runs, 0).image!;
		const floating = at(expectParagraph(model.blocks[3]).runs, 0).image!;
		expect(inline).toMatchObject({
			unsupported: 'SmartArt',
			altText: 'Organisation chart with assistants',
			widthPx: 867,
			heightPx: 533,
		});
		expect(inline.anchored).toBeUndefined();
		expect(floating).toMatchObject({
			unsupported: 'SmartArt',
			anchored: true,
			altText: 'Second organisation chart',
			title: 'Floating diagram',
			placement: { wrap: 'topAndBottom', align: 'center' },
		});
	});

	it('reports honestly: a cached snapshot, not recomputed Word layout', async () => {
		const { model } = await loadDocx(fixture());
		const warning = model.warnings.find((text) => text.includes('SmartArt'));
		expect(warning).toContain('2 SmartArt diagrams');
		expect(warning).toContain('Word layout is not recomputed');
		expect(warning).toContain('not editable');
		expect(model.warnings.some((text) => text.includes('rendered as a labeled placeholder'))).toBe(
			false,
		);
		expect(diagramsOf(model)[0]!.notice).toContain('not recomputed');
	});

	it('resolves the cached drawing colours against the document theme', async () => {
		const { model } = await loadDocx(fixture());
		const shape = diagramsOf(model)[0]!.drawing!.shapes.find((candidate) => candidate.text)!;
		const fill = shape.fill;
		expect(fill?.kind).toBe('solid');
		if (fill?.kind !== 'solid') return;
		const resolved = resolveDiagramColor(fill.color, model.theme);
		expect(resolved?.hex).toMatch(/^#[0-9A-F]{6}$/);
		expect(resolved?.hex).toBe(`#${model.theme!.colors.accent1}`);
		expect(resolveDiagramColor(fill.color, undefined)).toBeUndefined();
	});
});

describe('SmartArt preservation', () => {
	it('returns the package unchanged when nothing was edited', async () => {
		const original = fixture();
		const loaded = await loadDocx(original);
		expect(await loaded.save(loaded.model)).toEqual(original);
	});

	it('keeps every diagram part byte for byte and the drawing markup when other text is edited', async () => {
		const original = fixture();
		const loaded = await loadDocx(original);
		const edited = structuredClone(loaded.model);
		expectParagraph(edited.blocks[0]).runs[0]!.text = 'Edited intro text.';
		const saved = await loaded.save(edited);
		expect(await parts(saved, DIAGRAM_PARTS)).toEqual(await parts(original, DIAGRAM_PARTS));

		const zip = await JSZip.loadAsync(saved);
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).toContain('Edited intro text.');
		expect(document).toContain('uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"');
		expect(document.match(/<dgm:relIds /gu)).toHaveLength(2);
		const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
		expect(rels.match(/diagram(Data|Layout|QuickStyle|Colors|Drawing)"/gu)).toHaveLength(10);

		// Reopening yields the same diagram models (a round trip).
		const reopened = await loadDocx(saved);
		expect(JSON.parse(JSON.stringify(diagramsOf(reopened.model)))).toEqual(
			JSON.parse(JSON.stringify(diagramsOf(loaded.model))),
		);
	});

	it('keeps the other diagram intact when one diagram paragraph is deleted', async () => {
		const original = fixture();
		const loaded = await loadDocx(original);
		const edited = structuredClone(loaded.model);
		edited.blocks.splice(1, 1);
		const saved = await loaded.save(edited);
		const reopened = await loadDocx(saved);
		const remaining = diagramsOf(reopened.model);
		expect(remaining).toHaveLength(1);
		expect(remaining[0]).toMatchObject({ placement: 'anchor', rendering: 'cached-drawing' });
		const kept = DIAGRAM_PARTS.filter((name) => name.endsWith('2.xml'));
		expect(kept).toHaveLength(5);
		expect(await parts(saved, kept)).toEqual(await parts(original, kept));
	});
});

describe('SmartArt fallbacks', () => {
	it('falls back to a labelled placeholder with the node text when the package has no cached drawing', async () => {
		const loaded = await loadDocx(await withoutDiagramDrawing(fixture(), 1));
		const [first, second] = diagramsOf(loaded.model) as [DocxDiagram, DocxDiagram];
		expect(first.rendering).toBe('placeholder');
		expect(first.drawing).toBeUndefined();
		expect(first.notice).toContain('no usable cached drawing');
		expect(first.issues.map((issue) => issue.code)).toContain('DIAGRAM_DRAWING_ABSENT');
		expect(first.nodes.some((node) => node.text.length > 0)).toBe(true);
		// The other diagram keeps its own drawing and is not given this one.
		expect(second.rendering).toBe('cached-drawing');
		expect(second.parts.drawing?.partName).toBe('word/diagrams/drawing2.xml');
		const warnings = loaded.model.warnings.join('\n');
		expect(warnings).toContain('1 SmartArt diagram can be drawn');
		expect(warnings).toContain('SmartArt content is rendered as a labeled placeholder');
	});

	it('never borrows another diagram drawing when the data model names none', async () => {
		const zip = await JSZip.loadAsync(fixture());
		const data = await zip.file('word/diagrams/data1.xml')!.async('string');
		zip.file('word/diagrams/data1.xml', data.replace(/<dgm:extLst>[\s\S]*<\/dgm:extLst>/u, ''));
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const first = diagramsOf(loaded.model)[0]!;
		expect(first.rendering).toBe('placeholder');
		expect(first.parts.drawing).toBeUndefined();
	});

	it('reports a missing or unreadable part and still shows what it can', async () => {
		const zip = await JSZip.loadAsync(fixture());
		zip.remove('word/diagrams/layout1.xml');
		zip.file('word/diagrams/colors2.xml', '<not xml');
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const [first, second] = diagramsOf(loaded.model) as [DocxDiagram, DocxDiagram];
		expect(first.issues.map((issue) => issue.code)).toEqual(['DIAGRAM_PART_MISSING']);
		expect(first.layout).toBeUndefined();
		expect(first.rendering).toBe('cached-drawing');
		expect(second.issues.map((issue) => issue.code)).toEqual(['DIAGRAM_PART_UNREADABLE']);
		expect(
			loaded.model.warnings.some((text) => text.includes('2 problems found reading SmartArt')),
		).toBe(true);
	});

	it('reports a graphic without dgm:relIds', async () => {
		const zip = await JSZip.loadAsync(fixture());
		const document = await zip.file('word/document.xml')!.async('string');
		zip.file('word/document.xml', document.replace(/<dgm:relIds [^>]*\/>/u, ''));
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const first = diagramsOf(loaded.model)[0]!;
		expect(first.issues.map((issue) => issue.code)).toContain('DIAGRAM_RELIDS_MISSING');
		expect(first.rendering).toBe('placeholder');
	});
});
