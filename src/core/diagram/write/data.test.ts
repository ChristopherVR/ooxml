import { describe, expect, it } from 'vitest';

import { parseDiagramDataModel } from '../data-model';
import { writeFabricatedDiagramData } from './data';

const ids = {
	layoutUniqueId: 'urn:microsoft.com/office/officeart/2005/8/layout/default',
	layoutCategory: 'list',
	quickStyleUniqueId: 'urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1',
	colorsUniqueId: 'urn:microsoft.com/office/officeart/2005/8/colors/accent1_2',
	colorsCategory: 'accent1',
};

describe('writeFabricatedDiagramData', () => {
	it('writes a data model the neutral reader reads back as the same tree', () => {
		let counter = 0;
		const createModelId = () => `{00000000-0000-4000-8000-${String(++counter).padStart(12, '0')}}`;
		const xml = writeFabricatedDiagramData(
			{
				nodes: [
					{ id: 'a', text: 'Alpha & co' },
					{ id: 'b', text: 'Beta' },
					{ id: 'b1', text: '', parentId: 'b', nodeType: 'asst' },
				],
			},
			ids,
			{ createModelId, drawingRelId: 'rId7' },
		);
		const model = parseDiagramDataModel(xml);
		expect(model.issues).toEqual([]);
		expect(model.drawingRelId).toBe('rId7');
		const doc = model.points.find((point) => point.type === 'doc')!;
		expect(doc.layoutTypeId).toBe(ids.layoutUniqueId);
		expect(model.nodes.map((point) => [point.text, point.type])).toEqual([
			['Alpha & co', 'node'],
			['Beta', 'node'],
			['', 'asst'],
		]);
		const [alpha, beta, child] = model.nodes;
		expect(model.parentById.get(alpha!.modelId)).toBe(doc.modelId);
		expect(model.parentById.get(child!.modelId)).toBe(beta!.modelId);
		expect(model.points.filter((point) => point.type === 'pres')).toHaveLength(3);
		expect(xml).toContain('phldrT="[Text]"');
	});
});
