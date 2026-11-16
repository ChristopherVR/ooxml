import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseDiagramDataModel } from '../data-model';
import { parseDiagramDrawing } from '../drawing';
import { buildGeneratedDiagramCache } from './cache';
import { writeEditedDiagramData } from './data-edit';

// Word laid this diagram out itself (Basic Block List, five nodes, 6858000 x 2000250 EMU).
async function wordParts() {
	const zip = await JSZip.loadAsync(
		readFileSync(path.join(import.meta.dirname, '../../docx/__fixtures__/smartart-word.docx')),
	);
	const part = (name: string) => zip.file(`word/diagrams/${name}1.xml`)!.async('string');
	const [dataXml, layoutXml, colorsXml, quickStyleXml, drawingXml] = await Promise.all(
		['data', 'layout', 'colors', 'quickStyle', 'drawing'].map(part),
	);
	const document = await zip.file('word/document.xml')!.async('string');
	const [, cx, cy] = /<wp:extent cx="(\d+)" cy="(\d+)"/u.exec(document)!;
	return {
		dataXml: dataXml!,
		drawingXml: drawingXml!,
		input: {
			extentEmu: { width: Number(cx), height: Number(cy) },
			source: { layoutXml: layoutXml!, colorsXml: colorsXml!, quickStyleXml: quickStyleXml! },
			layoutId: /uniqueId="([^"]+)"/u.exec(layoutXml!)![1]!,
		},
	};
}

const presOf = (xml: string) =>
	parseDiagramDataModel(xml).connections.filter((edge) => edge.type === 'presOf');

describe('regenerating a source cache Office laid out', () => {
	it("binds every shape to Office's presentation point and matches its placement", async () => {
		const { dataXml, drawingXml, input } = await wordParts();
		const cache = buildGeneratedDiagramCache(dataXml, input)!;
		const office = parseDiagramDrawing(drawingXml).shapes;
		const regenerated = parseDiagramDrawing(cache.drawingXml).shapes;
		expect(regenerated.map((shape) => shape.modelId)).toEqual(office.map((shape) => shape.modelId));
		for (const [index, shape] of regenerated.entries()) {
			const frame = office[index]!.frame;
			for (const key of ['x', 'y', 'width', 'height'] as const)
				expect(Math.abs(shape.frame[key] - frame[key])).toBeLessThanOrEqual(2);
		}
		// No parallel presentation points: the data part keeps Office's points and connections.
		expect(parseDiagramDataModel(cache.dataXml).points.length).toBe(
			parseDiagramDataModel(dataXml).points.length,
		);
		expect(presOf(cache.dataXml)).toHaveLength(presOf(dataXml).length);
		expect(cache.dataXml).toContain('reusedPresentation="5"');
	});

	it('keeps the remaining points after an edit and adds one only for the new node', async () => {
		const { dataXml, drawingXml, input } = await wordParts();
		const model = parseDiagramDataModel(dataXml);
		const nodes = model.nodes.map((node) => ({
			id: node.modelId,
			text: node.text === 'Beta' ? 'Beta 2' : node.text,
		}));
		const added = '{00000000-0000-4000-8000-000000000001}';
		nodes.push({ id: added, text: 'Zeta' });
		const edited = writeEditedDiagramData(dataXml, nodes, { dropDrawingLink: true });
		const cache = buildGeneratedDiagramCache(edited, input)!;
		const ids = parseDiagramDrawing(cache.drawingXml).shapes.map((shape) => shape.modelId);
		const office = new Set(parseDiagramDrawing(drawingXml).shapes.map((shape) => shape.modelId));
		expect(ids).toHaveLength(6);
		expect(ids.filter((id) => office.has(id))).toHaveLength(5);
		const presented = presOf(cache.dataXml).map((edge) => edge.sourceId);
		expect(presented.filter((id) => id === added)).toHaveLength(1);
		// Each node keeps one presentation point per layout role: no duplicates beside Office's.
		const regenerated = buildGeneratedDiagramCache(cache.dataXml, input)!;
		expect(presOf(regenerated.dataXml)).toHaveLength(presOf(cache.dataXml).length);
		expect(
			parseDiagramDrawing(regenerated.drawingXml).shapes.map((shape) => shape.modelId),
		).toEqual(ids.map((id) => (office.has(id) ? id : expect.any(String))));
	});
});
