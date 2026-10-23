import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseDiagramDataModel } from '../data-model';
import { writeEditedDiagramData, type DiagramDataEditNode } from './data-edit';
import { writeFabricatedDiagramData } from './data';

// A data part Word wrote (Basic Block List, five nodes; docx/__fixtures__/smartart-word.docx).
const wordData = async (): Promise<string> => {
	const bytes = readFileSync(
		path.join(import.meta.dirname, '../../docx/__fixtures__/smartart-word.docx'),
	);
	return (await JSZip.loadAsync(bytes)).file('word/diagrams/data1.xml')!.async('string');
};

const contentNodes = (xml: string): DiagramDataEditNode[] => {
	const model = parseDiagramDataModel(xml);
	return model.nodes.map((point) => ({
		id: point.modelId,
		text: point.text,
		parentId: model.parentById.get(point.modelId),
	}));
};

let serial = 0;
const ids = () => `{00000000-0000-4000-8000-${String(++serial).padStart(12, '0')}}`;

describe('writeEditedDiagramData', () => {
	it('rewrites only the edited text and keeps the rest of the part', async () => {
		const xml = await wordData();
		const nodes = contentNodes(xml);
		const edited = nodes.map((node) => (node.text === 'Beta' ? { ...node, text: 'Beta 2' } : node));
		const out = writeEditedDiagramData(xml, edited);
		const model = parseDiagramDataModel(out);
		expect(model.nodes.map((node) => node.text)).toEqual([
			'Alpha',
			'Beta 2',
			'Gamma',
			'Delta',
			'Epsilon',
		]);
		// Same points, connections and presentation tree: only one text body changed.
		expect(model.points).toHaveLength(parseDiagramDataModel(xml).points.length);
		expect(out.replace('Beta 2', 'Beta')).toBe(xml);
	});

	it('adds, removes and reorders nodes with their connections and transitions', async () => {
		const xml = await wordData();
		const [alpha, beta, gamma, delta, epsilon] = contentNodes(xml) as DiagramDataEditNode[] & {
			length: 5;
		};
		const before = parseDiagramDataModel(xml);
		const out = writeEditedDiagramData(
			xml,
			[gamma!, alpha!, { id: 'new-1', text: 'Zeta', parentId: alpha!.id }, delta!, epsilon!],
			{ createModelId: ids },
		);
		const model = parseDiagramDataModel(out);
		expect(model.nodes.map((node) => node.text).sort()).toEqual(
			['Alpha', 'Delta', 'Epsilon', 'Gamma', 'Zeta'].sort(),
		);
		// Beta's point, its transitions, its presentation points and their connections are gone.
		expect(model.points.some((point) => point.modelId === beta!.id)).toBe(false);
		expect(out).not.toContain(beta!.id);
		const zeta = model.nodes.find((node) => node.text === 'Zeta')!;
		expect(zeta.modelId).toMatch(/^\{0{8}-/u);
		expect(model.parentById.get(zeta.modelId)).toBe(alpha!.id);
		const order = (id: string) =>
			model.connections.find((c) => c.destId === id && (c.type ?? 'parOf') === 'parOf')?.srcOrd;
		expect([gamma, alpha, delta, epsilon].map((node) => order(node!.id))).toEqual([0, 1, 2, 3]);
		expect(order(zeta.modelId)).toBe(0);
		const transitions = model.points.filter((point) => point.connectionId);
		expect(transitions.length).toBe(before.points.filter((p) => p.connectionId).length);
		// Every presOf left points at a remaining node.
		for (const cxn of model.connections.filter((c) => c.type === 'presOf'))
			expect(model.points.some((point) => point.modelId === cxn.sourceId)).toBe(true);
	});

	it('drops the link to the cached drawing on request', async () => {
		const xml = await wordData();
		expect(xml).toContain('dsp:dataModelExt');
		const out = writeEditedDiagramData(xml, contentNodes(xml), { dropDrawingLink: true });
		expect(out).not.toContain('dataModelExt');
		expect(parseDiagramDataModel(out).drawingRelId).toBeUndefined();
	});

	it('fills a fabricated skeleton with nodes and no presentation points', () => {
		const skeleton = writeFabricatedDiagramData(
			{ nodes: [] },
			{
				layoutUniqueId: 'urn:microsoft.com/office/officeart/2005/8/layout/default',
				layoutCategory: 'list',
				quickStyleUniqueId: 'urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1',
				colorsUniqueId: 'urn:microsoft.com/office/officeart/2005/8/colors/accent1_2',
				colorsCategory: 'accent1',
			},
			{ createModelId: ids },
		);
		const out = writeEditedDiagramData(
			skeleton,
			[
				{ id: 'a', text: 'One' },
				{ id: 'b', text: '' },
				{ id: 'c', text: 'Two\nlines', parentId: 'a' },
			],
			{ createModelId: ids },
		);
		const model = parseDiagramDataModel(out);
		expect(model.nodes.map((node) => node.text)).toEqual(['One', '', 'Two\nlines']);
		expect(model.points.filter((point) => point.type === 'pres')).toHaveLength(0);
		expect(out).toContain('phldrT="[Text]"');
		const one = model.nodes[0]!.modelId;
		expect(model.parentById.get(model.nodes[2]!.modelId)).toBe(one);
	});

	it('rejects a part without a document point', () => {
		expect(() =>
			writeEditedDiagramData(
				'<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst/></dgm:dataModel>',
				[],
			),
		).toThrow(/document point/u);
	});
});
