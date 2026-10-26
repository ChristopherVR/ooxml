import { expect, it } from 'vitest';
import { buildGeneratedDiagramCache, withDiagramDrawingRelationship } from './cache';
import { writeInsertedDiagramData } from './inserted';
import {
	builtinLayoutIdForPreset,
	loadBuiltinSmartArtLayoutXml,
	BUILTIN_SMARTART_COLORS_XML,
	BUILTIN_SMARTART_QUICK_STYLE_XML,
} from '../builtin';
import { parseDiagramDataModel } from '../data-model';
import { parseDiagramDrawing } from '../drawing';
import { loadDiagram } from '../load';
import { generatedDrawingDigest } from '../cache-metadata';

const layoutId = builtinLayoutIdForPreset('basicBlockList')!;
const source = {
	layoutXml: loadBuiltinSmartArtLayoutXml(layoutId),
	colorsXml: BUILTIN_SMARTART_COLORS_XML,
	quickStyleXml: BUILTIN_SMARTART_QUICK_STYLE_XML,
};
const input = {
	source,
	layoutId,
	extentEmu: { width: 600 * 9525, height: 350 * 9525 },
	scheme: { accent1: '#4472C4', lt1: '#FFFFFF', tx1: '#000000' },
};
const data = () =>
	writeInsertedDiagramData({
		nodes: [
			{ id: 'a', text: 'Alpha' },
			{ id: 'b', text: 'Beta' },
		],
		layoutId,
		colorsId: 'colors',
		quickStyleId: 'style',
	});

it('binds every generated shape GUID to an existing semantic data node and preserves foreign XML', () => {
	const xml = data()
		.replace('<dgm:dataModel ', '<dgm:dataModel xmlns:x="urn:kept" ')
		.replace(
			'</dgm:dataModel>',
			'<dgm:extLst><a:ext uri="urn:kept"><x:future value="1"/></a:ext></dgm:extLst></dgm:dataModel>',
		);
	const cache = buildGeneratedDiagramCache(xml, input)!;
	expect(cache).toBeDefined();
	expect(cache.dataXml).toContain('<x:future value="1"');
	const model = parseDiagramDataModel(cache.dataXml);
	const drawing = parseDiagramDrawing(cache.drawingXml);
	expect(model.generatedCacheDigest).toBe(generatedDrawingDigest(cache.drawingXml));
	const contentIds = new Set(model.nodes.map((node) => node.modelId));
	for (const shape of drawing.shapes) {
		const presentation = model.points.find((point) => point.modelId === shape.modelId);
		expect(presentation?.type).toBe('pres');
		expect(contentIds.has(presentation!.presAssocId!)).toBe(true);
		expect(
			model.connections.some(
				(edge) =>
					edge.type === 'presOf' &&
					edge.destId === shape.modelId &&
					edge.sourceId === presentation?.presAssocId,
			),
		).toBe(true);
	}
	expect(JSON.stringify(drawing.shapes.map((shape) => shape.fill))).toContain('scheme');
});

it('labels only the exact generated drawing bytes as generated, clearing stale markers after native rewrites', async () => {
	const cache = buildGeneratedDiagramCache(data(), input)!;
	const dataXml = withDiagramDrawingRelationship(cache.dataXml, 'draw');
	async function read(drawingXml: string) {
		const parts: Record<string, string> = {
			data: dataXml,
			layout: source.layoutXml,
			colors: source.colorsXml,
			style: source.quickStyleXml,
			drawing: drawingXml,
		};
		return loadDiagram(
			'host',
			{ dataRelId: 'dm', layoutRelId: 'lo', colorsRelId: 'cs', styleRelId: 'qs' },
			{
				hostRelationships: () =>
					new Map([
						['dm', 'data'],
						['lo', 'layout'],
						['cs', 'colors'],
						['qs', 'style'],
						['draw', 'drawing'],
					]),
				resolvePath: (_base, target) => target,
				readRelationshipEntries: async () => [],
				readText: async (part) => parts[part],
			},
		);
	}
	expect((await read(cache.drawingXml)).generatedCache).toBe(true);
	expect((await read(`${cache.drawingXml}\n`)).generatedCache).toBeUndefined();
});

it('declines unknown algorithms and spatial scenes without fabricating an Office cache', () => {
	expect(
		buildGeneratedDiagramCache(data(), {
			...input,
			source: { ...source, layoutXml: '<invalid/>' },
		}),
	).toBeUndefined();
	expect(
		buildGeneratedDiagramCache(data(), {
			...input,
			source: {
				...source,
				quickStyleXml: source.quickStyleXml.replaceAll(
					'orthographicFront',
					'perspectiveHeroicExtremeLeft',
				),
			},
		}),
	).toBeUndefined();
});

it('replaces only its own prior presentation points when regenerating', () => {
	const first = buildGeneratedDiagramCache(data(), input)!;
	const second = buildGeneratedDiagramCache(first.dataXml, input)!;
	expect(parseDiagramDataModel(second.dataXml).points.length).toBe(
		parseDiagramDataModel(first.dataXml).points.length,
	);
	expect(parseDiagramDataModel(second.dataXml).nodes.map((node) => node.text)).toEqual([
		'Alpha',
		'Beta',
	]);
});
