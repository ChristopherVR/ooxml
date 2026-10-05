import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRelationships, relationshipsPartFor, resolvePartPath } from '../opc/index.js';
import { parseXml } from '../xml/index.js';
import {
	attributeReader,
	collectTransitionText,
	loadDiagram,
	parseConnectionAttributes,
	parseCustomLayoutAttributes,
	parseDiagramColors,
	parseDiagramDataModel,
	parseDiagramDrawing,
	parseDiagramLayoutSummary,
	parseDiagramQuickStyle,
	parseRelationshipIdAttributes,
	resolveDiagramDrawingPart,
	resolveDiagramLayoutCategory,
	resolveDrawingColor,
	type DiagramPackageHost,
} from './index.js';

const fixtures = path.join(import.meta.dirname, '../pptx/__tests__/fixtures');
const deck = async (...segments: string[]) =>
	JSZip.loadAsync(readFileSync(path.join(fixtures, ...segments)));

/** A diagram package host over a PowerPoint package: the slide is the host part. */
function zipHost(
	zip: JSZip,
	relationships = new Map<string, Map<string, string>>(),
): DiagramPackageHost {
	return {
		hostRelationships: (part) => relationships.get(part),
		resolvePath: resolvePartPath,
		readText: async (part) => zip.file(part)?.async('string'),
		readRelationshipEntries: async (relsPart) => {
			const xml = await zip.file(relsPart)?.async('string');
			return xml ? [...parseRelationships(xml).values()] : undefined;
		},
	};
}

async function loadSlideDiagram(zip: JSZip, slide: string) {
	const rels = parseRelationships(await zip.file(relationshipsPartFor(slide))!.async('string'));
	const host = zipHost(
		zip,
		new Map([[slide, new Map([...rels].map(([id, rel]) => [id, rel.target]))]]),
	);
	const slideXml = parseXml(await zip.file(slide)!.async('string'));
	const relIds = slideXml.getElementsByTagNameNS(
		'http://schemas.openxmlformats.org/drawingml/2006/diagram',
		'relIds',
	)[0]!;
	return loadDiagram(slide, parseRelationshipIdAttributes(attributeReader(relIds)), host);
}

describe('attribute parsers', () => {
	const reader = (values: Record<string, string>) => (name: string) => values[name];

	it('parses a connection and trims its optional attributes', () => {
		expect(
			parseConnectionAttributes(
				reader({
					srcId: ' a ',
					destId: 'b',
					modelId: 'm',
					srcOrd: '2',
					destOrd: 'x',
					type: 'presOf',
				}),
			),
		).toEqual({ sourceId: 'a', destId: 'b', modelId: 'm', type: 'presOf', srcOrd: 2 });
		expect(parseConnectionAttributes(reader({ srcId: 'a' }))).toBeUndefined();
	});

	it('converts custom layout units: 60000ths of a degree, 100000ths ratios, booleans', () => {
		expect(
			parseCustomLayoutAttributes(
				reader({ custAng: '5400000', custScaleX: '150000', custFlipHor: '1', custFlipVert: '0' }),
			),
		).toEqual({ angle: 90, scaleX: 1.5, flipHorizontal: true, flipVertical: false });
		expect(parseCustomLayoutAttributes(reader({}))).toBeUndefined();
	});

	it('reads relIds by local name whatever the prefix', () => {
		const element = parseXml(
			'<x:relIds xmlns:x="urn:x" xmlns:q="urn:q" q:dm="rId1" q:lo="rId2" cs="rId4"/>',
		).documentElement;
		expect(parseRelationshipIdAttributes(attributeReader(element))).toEqual({
			dataRelId: 'rId1',
			layoutRelId: 'rId2',
			colorsRelId: 'rId4',
		});
	});

	it('resolves the layout family from category first, then unique id', () => {
		expect(resolveDiagramLayoutCategory('urn:x/layout/default', ['process'])).toBe('process');
		expect(resolveDiagramLayoutCategory('urn:x/layout/orgChart1', ['custom'])).toBe('hierarchy');
		expect(resolveDiagramLayoutCategory('urn:x/layout/opaque', [])).toBeUndefined();
	});

	it('collects transition text only for non-empty parTrans/sibTrans points', () => {
		const points = [
			{ type: 'parTrans', id: 'p', text: ' yes ' },
			{ type: 'sibTrans', id: 's', text: '' },
			{ type: 'node', id: 'n', text: 'content' },
		];
		expect(
			collectTransitionText(
				points,
				(point) => ({ type: point.type, modelId: point.id }),
				(point) => point.text,
			),
		).toEqual(new Map([['p', 'yes']]));
	});
});

describe('drawing colour', () => {
	const theme = { scheme: (name: string) => ({ accent1: '#4472C4', lt1: '#FFFFFF' })[name] };
	const colorOf = (xml: string) => {
		const color = parseDiagramColors(
			`<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="u"><dgm:styleLbl name="n"><dgm:fillClrLst>${xml}</dgm:fillClrLst></dgm:styleLbl></dgm:colorsDef>`,
		).labels[0]!.fill.colors[0]!;
		return color;
	};

	it('resolves srgb, scheme and transforms', () => {
		expect(resolveDrawingColor(colorOf('<a:srgbClr val="ff0000"/>'))).toEqual({
			hex: '#FF0000',
			alpha: 1,
			unapplied: [],
		});
		expect(resolveDrawingColor(colorOf('<a:schemeClr val="accent1"/>'), theme)?.hex).toBe(
			'#4472C4',
		);
		const shaded = resolveDrawingColor(
			colorOf('<a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr>'),
			theme,
		);
		expect(shaded?.hex).not.toBe('#4472C4');
		expect(
			resolveDrawingColor(
				colorOf('<a:schemeClr val="lt1"><a:alpha val="40000"/></a:schemeClr>'),
				theme,
			)?.alpha,
		).toBeCloseTo(0.4);
	});

	it('is undefined for an unknown theme slot and lists transforms it does not apply', () => {
		expect(resolveDrawingColor(colorOf('<a:schemeClr val="accent9"/>'), theme)).toBeUndefined();
		expect(
			resolveDrawingColor(colorOf('<a:srgbClr val="000000"><a:redMod val="50000"/></a:srgbClr>'))
				?.unapplied,
		).toEqual(['redMod']);
	});
});

describe('definition and data parts of a PowerPoint SmartArt', () => {
	it('parses the colours, quick style and layout of an org chart', async () => {
		const zip = await deck('corpus', 'smartart-orgchart-assistants.pptx');
		const colors = parseDiagramColors(await zip.file('ppt/diagrams/colors1.xml')!.async('string'));
		expect(colors.uniqueId).toBe('urn:microsoft.com/office/officeart/2005/8/colors/accent1_2');
		expect(colors.categories[0]).toEqual({ type: 'accent1', priority: 11200 });
		const node1 = colors.labels.find((label) => label.name === 'node1')!;
		expect(node1.fill).toEqual({
			method: 'repeat',
			colors: [{ kind: 'scheme', value: 'accent1', transforms: [] }],
		});
		const style = parseDiagramQuickStyle(
			await zip.file('ppt/diagrams/quickStyle1.xml')!.async('string'),
		);
		expect(style.uniqueId).toBe('urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1');
		expect(style.labels.find((label) => label.name === 'node1')?.line?.index).toBe(2);
		expect(style.labels.find((label) => label.name === 'node1')?.font?.fontIndex).toBe('minor');
		expect(style.has3d).toBe(false);
		const layout = parseDiagramLayoutSummary(
			await zip.file('ppt/diagrams/layout1.xml')!.async('string'),
		);
		expect(layout.family).toBe('hierarchy');
		expect(layout.uniqueId).toContain('orgChart1');
		expect(layout.layoutNodeCount).toBeGreaterThan(3);
	});

	it('parses the data model: content points, parent edges, drawing relationship', async () => {
		const zip = await deck('corpus', 'smartart-orgchart-assistants.pptx');
		const model = parseDiagramDataModel(await zip.file('ppt/diagrams/data1.xml')!.async('string'));
		expect(model.drawingRelId).toBe('rId6');
		expect(model.nodes.length).toBeGreaterThan(2);
		expect(model.nodes.every((node) => !['doc', 'pres'].includes(node.type))).toBe(true);
		expect(model.issues).toEqual([]);
		const docPoint = model.points.find((point) => point.type === 'doc')!;
		expect(docPoint.layoutTypeId).toContain('orgChart1');
		const child = [...model.parentById.keys()][0]!;
		expect(model.points.some((point) => point.modelId === model.parentById.get(child))).toBe(true);
		// Every connection endpoint exists.
		expect(model.connections.length).toBeGreaterThan(0);
	});

	it('reports a data model with a missing endpoint instead of throwing', () => {
		const model = parseDiagramDataModel(
			'<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst><dgm:pt modelId="a"/><dgm:pt modelId="a"/><dgm:pt/></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="c" srcId="a" destId="zz"/><dgm:cxn modelId="d"/></dgm:cxnLst></dgm:dataModel>',
		);
		expect(model.issues.map((issue) => issue.code)).toEqual([
			'POINT_ID_DUPLICATE',
			'POINT_ID_REQUIRED',
			'CONNECTION_ENDPOINT_MISSING',
			'CONNECTION_ATTRIBUTE_REQUIRED',
		]);
	});
});

describe('cached drawing', () => {
	it('reads a PowerPoint-fabricated list: preset geometry, solid fill, line, text, text frame', async () => {
		const zip = await deck('e2e', 'smartart-build-reveal.pptx');
		const drawing = parseDiagramDrawing(
			await zip.file('ppt/diagrams/drawing1.xml')!.async('string'),
		);
		expect(drawing.issues).toEqual([]);
		expect(drawing.shapes.map((shape) => shape.text?.text)).toEqual(['Alpha', 'Beta', 'Gamma']);
		const [alpha] = drawing.shapes;
		expect(alpha).toMatchObject({
			modelId: '{00000000-0000-4000-8000-000000000005}',
			frame: { x: 76200, y: 76200, width: 5562600, height: 1019175 },
			geometry: 'roundRect',
			fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4' } },
			line: { widthEmu: 12700 },
			textFrame: { x: 76200, y: 76200, width: 5562600, height: 1019175 },
			has3d: false,
		});
		expect(alpha!.text).toMatchObject({
			anchor: 'ctr',
			paragraphs: [{ align: 'ctr', runs: [{ text: 'Alpha', sizePt: 8.25, typeface: 'Calibri' }] }],
		});
		expect(alpha!.style?.fill).toEqual({
			index: 1,
			color: { kind: 'scheme', value: 'accent1', transforms: [] },
		});
	});

	it('reads connector shapes with custom geometry and no fill', async () => {
		const zip = await deck('corpus', 'smartart-orgchart-assistants.pptx');
		const drawing = parseDiagramDrawing(
			await zip.file('ppt/diagrams/drawing1.xml')!.async('string'),
		);
		const connector = drawing.shapes[0]!;
		expect(connector.geometry).toBe('custom');
		expect(connector.fill).toEqual({ kind: 'none' });
		expect(connector.paths?.[0]?.commands).toEqual([
			{ op: 'M', x: 253408, y: 0 },
			{ op: 'L', x: 253408, y: 1110172 },
			{ op: 'L', x: 0, y: 1110172 },
		]);
		const accent = connector.line?.fill;
		expect(accent).toMatchObject({ kind: 'solid', color: { value: 'accent1' } });
		expect(drawing.shapes.some((shape) => shape.text)).toBe(true);
	});

	it('flags 3D content in a bevel/scene quick style deck', async () => {
		const zip = await deck('e2e', 'three-d-parity', 'three-d-smartart.pptx');
		const names = Object.keys(zip.files).filter((name) =>
			/ppt\/diagrams\/drawing\d+\.xml$/.test(name),
		);
		expect(names.length).toBeGreaterThan(50);
		let withShapes = 0;
		let with3d = 0;
		for (const name of names) {
			const drawing = parseDiagramDrawing(await zip.file(name)!.async('string'), name);
			if (drawing.shapes.length > 0) withShapes++;
			if (drawing.shapes.some((shape) => shape.has3d)) with3d++;
			expect(drawing.issues.filter((issue) => issue.code === 'DIAGRAM_DRAWING_EMPTY')).toEqual([]);
		}
		expect(withShapes).toBe(names.length);
		expect(with3d).toBeGreaterThan(0);
	});

	it('reports an empty drawing and skips pictures honestly', () => {
		expect(
			parseDiagramDrawing(
				'<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"/>',
			).issues[0]?.code,
		).toBe('DIAGRAM_DRAWING_EMPTY');
		const withPicture = parseDiagramDrawing(
			'<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:spTree><dsp:pic/></dsp:spTree></dsp:drawing>',
		);
		expect(withPicture.issues[0]?.code).toBe('DIAGRAM_DRAWING_PICTURE_SKIPPED');
	});

	it('maps nested group shapes through the group transform', () => {
		const drawing = parseDiagramDrawing(
			`<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree>
			<dsp:grpSp><dsp:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="2000" cy="2000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></dsp:grpSpPr>
			<dsp:sp modelId="g"><dsp:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst><a:gd name="adj" fmla="val 1234"/></a:avLst></a:prstGeom></dsp:spPr></dsp:sp></dsp:grpSp></dsp:spTree></dsp:drawing>`,
		);
		expect(drawing.shapes[0]).toMatchObject({
			frame: { x: 1200, y: 2200, width: 1000, height: 1000 },
			geometry: 'ellipse',
			adjustments: { adj: 1234 },
		});
	});
});

describe('resolveDiagramDrawingPart', () => {
	const hostFor = (rels: Record<string, string>, partRels?: string): DiagramPackageHost => ({
		hostRelationships: () => new Map(Object.entries(rels)),
		resolvePath: resolvePartPath,
		readText: async () => undefined,
		readRelationshipEntries: async () =>
			partRels ? [{ id: 'rId9', type: 'x/diagramDrawing', target: partRels }] : undefined,
	});

	it('prefers the dataModelExt relationship id on the host', async () => {
		expect(
			await resolveDiagramDrawingPart(
				'ppt/slides/slide1.xml',
				'rId2',
				'rId6',
				hostFor({ rId2: '../diagrams/data1.xml', rId6: '../diagrams/drawing1.xml' }),
			),
		).toEqual({ relId: 'rId6', path: 'ppt/diagrams/drawing1.xml' });
	});

	it('infers a drawing relationship by its target when the extension is missing', async () => {
		expect(
			await resolveDiagramDrawingPart(
				'ppt/slides/slide1.xml',
				'rId2',
				'',
				hostFor({ rId2: '../diagrams/data1.xml', rId7: '../diagrams/drawing3.xml' }),
			),
		).toEqual({
			relId: 'rId7',
			path: 'ppt/diagrams/drawing3.xml',
		});
	});

	it('falls back to the data part relationships, and to nothing', async () => {
		expect(
			await resolveDiagramDrawingPart(
				'ppt/slides/slide1.xml',
				'rId2',
				'',
				hostFor({ rId2: '../diagrams/data1.xml' }, 'drawing1.xml'),
			),
		).toEqual({ relId: 'rId9', path: 'ppt/diagrams/drawing1.xml' });
		expect(
			await resolveDiagramDrawingPart(
				'ppt/slides/slide1.xml',
				'rId2',
				'',
				hostFor({ rId2: '../diagrams/data1.xml' }),
			),
		).toBeUndefined();
		expect(
			await resolveDiagramDrawingPart('ppt/slides/slide1.xml', '', 'rId6', hostFor({})),
		).toBeUndefined();
	});
});

describe('loadDiagram', () => {
	it('loads every part of a real SmartArt slide', async () => {
		const zip = await deck('corpus', 'smartart-orgchart-assistants.pptx');
		const diagram = await loadSlideDiagram(zip, 'ppt/slides/slide1.xml');
		expect(diagram.paths).toEqual({
			data: 'ppt/diagrams/data1.xml',
			layout: 'ppt/diagrams/layout1.xml',
			quickStyle: 'ppt/diagrams/quickStyle1.xml',
			colors: 'ppt/diagrams/colors1.xml',
			drawing: 'ppt/diagrams/drawing1.xml',
		});
		expect(diagram.drawingRelId).toBe('rId6');
		expect(diagram.issues).toEqual([]);
		expect(diagram.layout?.family).toBe('hierarchy');
		expect(diagram.drawing!.shapes.length).toBe(9);
		expect(diagram.data!.nodes.length).toBeGreaterThan(0);
	});

	it('reports missing relationships and unreadable parts and still loads the rest', async () => {
		const zip = await deck('corpus', 'smartart-orgchart-assistants.pptx');
		zip.file('ppt/diagrams/colors1.xml', '<not xml');
		const diagram = await loadSlideDiagram(zip, 'ppt/slides/slide1.xml');
		expect(diagram.colors).toBeUndefined();
		expect(diagram.issues.map((issue) => issue.code)).toEqual(['DIAGRAM_PART_UNREADABLE']);
		expect(diagram.data).toBeDefined();
		const missing = await loadDiagram(
			'ppt/slides/slide1.xml',
			{ dataRelId: 'rId2', layoutRelId: 'rIdX' },
			zipHost(zip),
		);
		expect(missing.issues.map((issue) => issue.code)).toContain('DIAGRAM_RELATIONSHIP_MISSING');
	});

	it('says so when no cached drawing exists', async () => {
		const zip = await deck('corpus', 'smartart-orgchart-assistants.pptx');
		zip.remove('ppt/diagrams/drawing1.xml');
		const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
		zip.file(
			'ppt/slides/_rels/slide1.xml.rels',
			rels.replace(/<Relationship [^>]*drawing1\.xml"[^>]*\/>/u, ''),
		);
		const diagram = await loadSlideDiagram(zip, 'ppt/slides/slide1.xml');
		expect(diagram.drawing).toBeUndefined();
		expect(diagram.issues.map((issue) => issue.code)).toContain('DIAGRAM_DRAWING_ABSENT');
	});
});
