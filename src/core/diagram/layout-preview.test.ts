import { describe, expect, it } from 'vitest';
import {
	builtinLayoutIdForPreset,
	loadBuiltinSmartArtLayoutXml,
	listBuiltinSmartArtLayouts,
} from './builtin';
import { diagramLayoutData, previewDiagramLayout } from './layout-preview';
import { runEngineLayout } from './engine';

const input = {
	extentEmu: { width: 600 * 9525, height: 350 * 9525 },
	nodes: [
		{ id: 'a', text: 'Plan' },
		{ id: 'b', text: 'Build' },
		{ id: 'c', text: 'Ship' },
	],
};

describe('uncached diagram previews', () => {
	it('uses the same ordered layout engine rather than a family approximation', () => {
		const layoutXml = loadBuiltinSmartArtLayoutXml(
			builtinLayoutIdForPreset('basicChevronProcess')!,
		);
		const source = { layoutXml };
		const preview = previewDiagramLayout({ ...input, source });
		expect(preview.mode).toBe('interpreted');
		const data = diagramLayoutData({ ...input, source });
		const expected = runEngineLayout(
			data,
			{ width: 600, height: 350 },
			data.nodes,
			Array(6).fill('#4472c4'),
			'flat',
		)!;
		expect(preview.drawing!.shapes.map((shape) => shape.frame)).toEqual(
			expected.nodes.map((node) => {
				if (node.kind !== 'rect') throw new Error('Engine result must use rect carriers');
				return {
					x: Math.round(node.x * 9525),
					y: Math.round(node.y * 9525),
					width: Math.round(node.width * 9525),
					height: Math.round(node.height * 9525),
				};
			}),
		);
		expect(preview.drawing!.shapes.some((shape) => shape.geometry === 'chevron')).toBe(true);
	});

	it('retains definitions, layout variables and transition metadata while current edits win', () => {
		const dataXml = `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst><dgm:pt modelId="a"><dgm:prSet><dgm:presLayoutVars><dgm:dir val="rev"/><dgm:hierBranch val="hang"/><dgm:orgChart val="1"/><dgm:chMax val="4"/></dgm:presLayoutVars></dgm:prSet></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="old" type="parOf" srcId="doc" destId="a" srcOrd="99" parTransId="t"/></dgm:cxnLst></dgm:dataModel>`;
		const source = {
			dataXml,
			layoutXml: '<original ordered constraints/>',
			quickStyleXml: '<style/>',
		};
		const data = diagramLayoutData({ ...input, nodes: [{ id: 'a', text: 'Edited' }], source });
		expect(data.layoutDefinition?.rawXmlText).toBe(source.layoutXml);
		expect(data.presLayoutVars).toMatchObject({
			direction: 'rev',
			hierarchyBranch: 'hang',
			orgChart: true,
			childMax: 4,
		});
		expect(data.nodes[0]?.text).toBe('Edited');
		expect(data.connections?.[0]).toMatchObject({ srcOrd: 0, parentTransitionId: 't' });
		expect(JSON.parse(JSON.stringify(source))).toEqual(source);
	});

	it('reports fallback for missing, invalid and unsupported definitions', () => {
		for (const layoutXml of [
			undefined,
			'<broken',
			'<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:layoutNode><dgm:alg type="unknown"/></dgm:layoutNode></dgm:layoutDef>',
		]) {
			const result = previewDiagramLayout({ ...input, source: layoutXml ? { layoutXml } : {} });
			expect(result.mode).toBe('approximate');
			expect(result.issues.some((issue) => issue.code === 'DIAGRAM_LAYOUT_APPROXIMATE')).toBe(true);
		}
	});

	it('runs the complete built-in corpus without silent empty or non-finite drawings', () => {
		let interpreted = 0;
		for (const entry of listBuiltinSmartArtLayouts()) {
			const preview = previewDiagramLayout({
				...input,
				source: { layoutXml: loadBuiltinSmartArtLayoutXml(entry.id) },
			});
			expect(preview.drawing?.shapes.length, entry.id).toBeGreaterThan(0);
			for (const shape of preview.drawing!.shapes)
				expect(Object.values(shape.frame).every(Number.isFinite), entry.id).toBe(true);
			if (preview.mode === 'interpreted') interpreted++;
			else
				expect(
					preview.issues.some((issue) => issue.code === 'DIAGRAM_LAYOUT_APPROXIMATE'),
					entry.id,
				).toBe(true);
		}
		expect(interpreted).toBeGreaterThan(100);
	});
});
