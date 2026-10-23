import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { diagramsIn, loadDocx, type DocxDiagram, type DocumentModel } from './index';
import { createDocxDiagram, editDocxDiagram, insertedDiagramPartNames } from './diagram-edit';
import { expectParagraph } from './test-support/access';

const fixture = (name: string) =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '__fixtures__', name)));
const texts = (diagram: DocxDiagram) => diagram.nodes.map((node) => node.text);

/** The model with the first diagram run's diagram replaced. */
function withFirstDiagram(model: DocumentModel, next: DocxDiagram): DocumentModel {
	const edited = structuredClone(model);
	for (const block of edited.blocks) {
		if (block.type !== 'paragraph') continue;
		const run = block.runs.find((candidate) => candidate.image?.diagram);
		if (run?.image) {
			run.image.diagram = next;
			return edited;
		}
	}
	throw new Error('no diagram');
}

let serial = 0;
const ids = () => `{00000000-0000-4000-8000-${String(++serial).padStart(12, '0')}}`;

describe('editing SmartArt in a Word document', () => {
	it('orders nodes as a tree and applies the shared node operations', async () => {
		const { model } = await loadDocx(fixture('smartart-word.docx'));
		const diagram = diagramsIn(model.blocks)[0]!;
		expect(texts(diagram)).toEqual(['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon']);
		const [alpha, beta] = diagram.nodes;
		const added = editDocxDiagram(
			diagram,
			{ kind: 'add', position: 'after', nodeId: alpha!.id, text: 'New' },
			{ createModelId: ids },
		);
		expect(added.changed).toBe(true);
		expect(texts(added.diagram)).toEqual(['Alpha', 'New', 'Beta', 'Gamma', 'Delta', 'Epsilon']);
		expect(added.nodeId).toBe(added.diagram.nodes[1]!.id);
		expect(added.diagram.drawing).toBeUndefined();
		expect(added.diagram.parts.drawing).toBeUndefined();
		expect(added.diagram.rendering).toBe('placeholder');

		const demoted = editDocxDiagram(added.diagram, { kind: 'demote', nodeId: beta!.id });
		expect(demoted.diagram.nodes.find((node) => node.id === beta!.id)?.parentId).toBe(added.nodeId);
		const promoted = editDocxDiagram(demoted.diagram, { kind: 'promote', nodeId: beta!.id });
		expect(texts(promoted.diagram)).toEqual(texts(added.diagram));
		expect(editDocxDiagram(diagram, { kind: 'demote', nodeId: alpha!.id }).changed).toBe(false);
		const moved = editDocxDiagram(diagram, { kind: 'move', nodeId: alpha!.id, direction: 1 });
		expect(texts(moved.diagram).slice(0, 2)).toEqual(['Beta', 'Alpha']);
		const renamed = editDocxDiagram(diagram, { kind: 'text', nodeId: beta!.id, text: 'B' });
		expect(texts(renamed.diagram)[1]).toBe('B');
		expect(renamed.diagram.nodes[1]!.paragraphs).toBeUndefined();
		expect(renamed.diagram.nodes[0]!.paragraphs).toBeDefined();
	});

	it('saves an edit to the data part, drops the stale drawing and reloads it', async () => {
		const original = fixture('smartart-word.docx');
		const loaded = await loadDocx(original);
		const diagram = diagramsIn(loaded.model.blocks)[0]!;
		const [, beta, gamma] = diagram.nodes;
		let next = editDocxDiagram(diagram, {
			kind: 'text',
			nodeId: beta!.id,
			text: 'Beta edited',
		}).diagram;
		next = editDocxDiagram(next, { kind: 'remove', nodeId: gamma!.id }).diagram;
		next = editDocxDiagram(
			next,
			{ kind: 'add', position: 'below', nodeId: beta!.id, text: 'Child' },
			{ createModelId: ids },
		).diagram;
		const saved = await loaded.save(withFirstDiagram(loaded.model, next));

		const zip = await JSZip.loadAsync(saved);
		expect(zip.file('word/diagrams/drawing1.xml')).toBeNull();
		const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
		expect(rels).not.toContain('diagramDrawing');
		const types = await zip.file('[Content_Types].xml')!.async('string');
		expect(types).not.toContain('drawing1.xml');
		const before = await JSZip.loadAsync(original);
		for (const part of ['layout1', 'colors1', 'quickStyle1'])
			expect(await zip.file(`word/diagrams/${part}.xml`)!.async('string')).toBe(
				await before.file(`word/diagrams/${part}.xml`)!.async('string'),
			);
		// The inline drawing markup is kept as it was.
		const documentXml = await zip.file('word/document.xml')!.async('string');
		const beforeXml = await before.file('word/document.xml')!.async('string');
		const graphic = (xml: string) => /<w:drawing>.*?<\/w:drawing>/su.exec(xml)?.[0];
		expect(graphic(documentXml)).toBe(graphic(beforeXml));

		const reloaded = diagramsIn((await loadDocx(saved)).model.blocks)[0]!;
		expect(texts(reloaded)).toEqual(['Alpha', 'Beta edited', 'Child', 'Delta', 'Epsilon']);
		expect(reloaded.nodes.find((node) => node.text === 'Child')?.parentId).toBe(beta!.id);
		expect(reloaded.rendering).toBe('placeholder');
		expect(reloaded.issues.map((issue) => issue.code)).toEqual(['DIAGRAM_DRAWING_ABSENT']);
	});

	it('keeps every byte of an unedited diagram when another one is edited', async () => {
		const original = fixture('smartart.docx');
		const loaded = await loadDocx(original);
		const [first] = diagramsIn(loaded.model.blocks) as [DocxDiagram];
		const node = first.nodes.find((candidate) => candidate.text)!;
		const next = editDocxDiagram(first, { kind: 'text', nodeId: node.id, text: 'Renamed' });
		const saved = await loaded.save(withFirstDiagram(loaded.model, next.diagram));
		const [after, before] = await Promise.all([saved, original].map((b) => JSZip.loadAsync(b)));
		for (const stem of ['data', 'layout', 'quickStyle', 'colors', 'drawing'])
			expect(await after!.file(`word/diagrams/${stem}2.xml`)!.async('string'), stem).toBe(
				await before!.file(`word/diagrams/${stem}2.xml`)!.async('string'),
			);
		expect(after!.file('word/diagrams/drawing1.xml')).toBeNull();
		const reopened = diagramsIn((await loadDocx(saved)).model.blocks);
		expect(reopened[1]!.rendering).toBe('cached-drawing');
		expect(reopened[0]!.nodes.some((candidate) => candidate.text === 'Renamed')).toBe(true);
	});

	it('returns the original bytes when an edit is undone', async () => {
		const original = fixture('smartart-word.docx');
		const loaded = await loadDocx(original);
		expect(Array.from(await loaded.save(structuredClone(loaded.model)))).toEqual(
			Array.from(original),
		);
	});

	it('inserts a new SmartArt with its four parts, relationships and content types', async () => {
		const loaded = await loadDocx(fixture('smartart-word.docx'));
		const diagram = createDocxDiagram({
			layoutId: 'urn:microsoft.com/office/officeart/2005/8/layout/default',
			family: 'list',
			colorsId: 'urn:microsoft.com/office/officeart/2005/8/colors/accent1_2',
			quickStyleId: 'urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1',
			items: ['One', 'Two'],
			widthPx: 400,
			heightPx: 240,
			partSuffix: '-test',
			createModelId: ids,
		});
		const model = structuredClone(loaded.model);
		expectParagraph(model.blocks[0]).runs.push({
			text: '',
			image: {
				relId: '',
				partName: '',
				contentType: 'application/octet-stream',
				widthPx: 400,
				heightPx: 240,
				unsupported: 'SmartArt',
				diagram,
			},
		});
		const names = insertedDiagramPartNames('-test');
		const definition = (root: string, id: string) => ({
			bytes: new TextEncoder().encode(
				`<dgm:${root} xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="${id}"/>`,
			),
			contentType: 'application/xml',
		});
		const saved = await loaded.save(
			model,
			new Map([
				[names.layout, definition('layoutDef', diagram.layout!.uniqueId)],
				[names.colors, definition('colorsDef', diagram.colorsId!)],
				[names.quickStyle, definition('styleDef', diagram.quickStyleId!)],
			]),
		);
		const zip = await JSZip.loadAsync(saved);
		const types = await zip.file('[Content_Types].xml')!.async('string');
		expect(types).toContain('PartName="/word/diagrams/data-test.xml"');
		expect(types).toContain('drawingml.diagramLayout+xml');
		const documentXml = await zip.file('word/document.xml')!.async('string');
		expect(documentXml.match(/<dgm:relIds /gu)).toHaveLength(2);
		const reloaded = diagramsIn((await loadDocx(saved)).model.blocks);
		expect(reloaded).toHaveLength(2);
		const inserted = reloaded.find((candidate) =>
			candidate.parts.data?.partName.includes('-test'),
		)!;
		expect(texts(inserted)).toEqual(['One', 'Two']);
		expect(inserted.layout?.uniqueId).toBe(diagram.layout!.uniqueId);
		expect(inserted.inserted).toBeUndefined();
		expect(inserted.issues.filter((issue) => issue.code !== 'DIAGRAM_DRAWING_ABSENT')).toEqual([]);
	});
});
