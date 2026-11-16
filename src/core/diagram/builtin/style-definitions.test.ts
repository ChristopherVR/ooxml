import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { diagramsIn, formatDocxDiagram, loadDocx } from '../../docx/index';
import { parseDiagramColors, parseDiagramQuickStyle } from '../definitions';
import {
	DIAGRAM_NATIVE_COLOR_ENTRIES,
	DIAGRAM_NATIVE_STYLE_ENTRIES,
	appliedDiagramNativeId,
	findDiagramNativeColor,
} from '../catalogue';
import { formatDiagramSource } from '../edit/format';
import {
	BUILTIN_SMARTART_STYLE_DEFINITIONS,
	builtinSmartArtFormatChange,
	loadBuiltinSmartArtStyleDefinitionXml,
} from './style-definitions';

const word = () =>
	new Uint8Array(
		readFileSync(new URL('../../docx/__fixtures__/smartart-word.docx', import.meta.url)),
	);

describe('built-in SmartArt colour and quick-style catalogue', () => {
	it('lists every Office definition once, in gallery order, matching the gallery metadata', () => {
		expect(DIAGRAM_NATIVE_COLOR_ENTRIES).toHaveLength(38);
		expect(DIAGRAM_NATIVE_STYLE_ENTRIES).toHaveLength(14);
		expect(BUILTIN_SMARTART_STYLE_DEFINITIONS.map((entry) => entry.id)).toEqual([
			...DIAGRAM_NATIVE_COLOR_ENTRIES.map((entry) => entry.uniqueId),
			...DIAGRAM_NATIVE_STYLE_ENTRIES.map((entry) => entry.uniqueId),
		]);
		expect(DIAGRAM_NATIVE_STYLE_ENTRIES.filter((entry) => entry.category === '3D')).toHaveLength(9);
		expect(
			findDiagramNativeColor('urn:microsoft.com/office/officeart/2005/8/colors/accent1_2'),
		).toMatchObject({ id: 'accent1_2', name: 'Colored Fill - Accent 1' });
		expect(appliedDiagramNativeId('colors', 'urn:ooxml-core/colors/monochromatic2')).toBe(
			'accent2_2',
		);
		expect(appliedDiagramNativeId('style', 'urn:ooxml-core/style/intense')).toBe('simple5');
	});

	it.each(DIAGRAM_NATIVE_COLOR_ENTRIES.map((entry) => [entry.id, entry] as const))(
		'parses colour definition %s',
		(_id, entry) => {
			const xml = loadBuiltinSmartArtStyleDefinitionXml('colors', entry.id);
			const parsed = parseDiagramColors(xml);
			expect(parsed.uniqueId).toBe(entry.uniqueId);
			const node = parsed.labels.find((label) => label.name === 'node1');
			expect(node?.fill.colors.length).toBeGreaterThan(0);
			expect(xml).not.toContain('<?xml');
		},
	);

	it.each(DIAGRAM_NATIVE_STYLE_ENTRIES.map((entry) => [entry.id, entry] as const))(
		'parses quick-style definition %s',
		(_id, entry) => {
			const parsed = parseDiagramQuickStyle(
				loadBuiltinSmartArtStyleDefinitionXml('style', entry.uniqueId),
			);
			expect(parsed.uniqueId).toBe(entry.uniqueId);
			const node = parsed.labels.find((label) => label.name === 'node1');
			expect(node?.fill?.index).toBe(entry.preview.fill);
			expect(node?.effect?.index).toBe(entry.preview.effect);
			expect(parsed.labels.some((label) => label.requires3d)).toBe(
				entry.category === '3D' || entry.preview.bevel === true,
			);
		},
	);

	it('replaces the definition part with the built-in one and records its uniqueId', () => {
		const change = builtinSmartArtFormatChange('style', '3d1');
		const next = formatDiagramSource({ quickStyleXml: '<dgm:styleDef/>' }, change);
		expect(next.quickStyleId).toBe('urn:microsoft.com/office/officeart/2005/8/quickstyle/3d1');
		expect(next.source.quickStyleXml).toBe(change.definitionXml);
		expect(next.source.quickStyleXml).toContain('<dgm:sp3d prstMaterial="plastic">');
		expect(() =>
			formatDiagramSource(undefined, { ...change, kind: 'colors', definitionXml: '<a/>' }),
		).toThrow();
		expect(() => builtinSmartArtFormatChange('colors', 'accent9_9')).toThrow();
	});
});

describe('applying every built-in definition to a Word SmartArt', () => {
	const entries = [
		...DIAGRAM_NATIVE_COLOR_ENTRIES.map((entry) => ['colors', entry] as const),
		...DIAGRAM_NATIVE_STYLE_ENTRIES.map((entry) => ['style', entry] as const),
	];
	it.each(entries.map(([kind, entry]) => [`${kind} ${entry.id}`, kind, entry] as const))(
		'saves and reopens %s',
		async (_label, kind, entry) => {
			const loaded = await loadDocx(word());
			const model = structuredClone(loaded.model);
			const run = model.blocks
				.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
				.find((candidate) => candidate.image?.diagram)!;
			const original = run.image!.diagram!;
			const change = builtinSmartArtFormatChange(kind, entry.id);
			run.image!.diagram = formatDocxDiagram(original, change);
			const reopened = diagramsIn((await loadDocx(await loaded.save(model))).model.blocks)[0]!;
			expect(kind === 'colors' ? reopened.colorsId : reopened.quickStyleId).toBe(entry.uniqueId);
			const saved =
				kind === 'colors' ? reopened.layoutSource?.colorsXml : reopened.layoutSource?.quickStyleXml;
			// Word's own copy (the fixture's default) keeps its declaration and line breaks.
			expect(saved?.replace(/^<\?xml[^>]*\?>\s*/u, '').replace(/>\s+</gu, '><')).toBe(
				change.definitionXml,
			);
			expect(reopened.nodes.map((node) => node.text)).toEqual(
				original.nodes.map((node) => node.text),
			);
			// The fixture already uses Word's defaults; choosing one again keeps Word's own cache.
			const unchanged = ['accent1_2', 'simple1'].includes(entry.id);
			expect(reopened.rendering).toBe(unchanged ? 'cached-drawing' : 'generated-drawing');
		},
	);

	it('writes the 3D a quick style gives its nodes into the generated cache', async () => {
		const loaded = await loadDocx(word());
		const model = structuredClone(loaded.model);
		const run = model.blocks
			.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
			.find((candidate) => candidate.image?.diagram)!;
		run.image!.diagram = formatDocxDiagram(
			run.image!.diagram!,
			builtinSmartArtFormatChange('style', '3d9'),
		);
		const zip = await JSZip.loadAsync(await loaded.save(model));
		const drawingName = Object.keys(zip.files).find((name) =>
			/^word\/diagrams\/drawing\d*\.xml$/u.test(name),
		)!;
		const drawing = await zip.file(drawingName)!.async('string');
		// Bird's Eye Scene: bevelled extruded nodes and extruded text, as Office's own cache has them.
		expect(drawing).toContain(
			'<a:sp3d extrusionH="152250" prstMaterial="matte"><a:bevelT w="165100" prst="coolSlant"/></a:sp3d></dsp:spPr>',
		);
		expect(drawing).toMatch(/<a:sp3d extrusionH="28000" prstMaterial="matte"\/><\/a:bodyPr>/u);
		expect(drawing).not.toContain('perspectiveRelaxed');
	});
});
