import type JSZip from 'jszip';
import {
	buildGeneratedDiagramCache,
	withDiagramDrawingRelationship,
	DIAGRAM_CONTENT_TYPES,
	DIAGRAM_RELATIONSHIP_TYPES,
} from '../diagram/index';
import type { DocxDiagram } from './diagram';
import { diagramsIn } from './diagram-document';
import { diagramColorTheme } from './diagram-theme';
import type { Block } from './model';
import type { ThemeCatalog } from './theme-model';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts';

const semantic = (diagram: DocxDiagram) =>
	JSON.stringify({
		nodes: diagram.nodes.map(({ id, text, parentId }) => ({ id, text, parentId })),
		layout: diagram.layout?.uniqueId,
		colors: diagram.colorsId,
		style: diagram.quickStyleId,
	});

/** Attach generated caches only to changed body diagrams; unsupported interpretations stay uncached. */
export async function applyGeneratedDiagramCaches(
	zip: JSZip,
	blocks: Block[],
	baseBlocks: Block[],
	theme: ThemeCatalog | undefined,
): Promise<void> {
	const base = new Map(
		diagramsIn(baseBlocks).map((diagram) => [diagram.parts.data?.partName, diagram]),
	);
	const done = new Set<string>();
	const resolver = diagramColorTheme(theme);
	const scheme: Record<string, string> = {};
	for (const key of [...Object.keys(theme?.colors ?? {}), 'bg1', 'tx1', 'bg2', 'tx2']) {
		const color = resolver.scheme(key);
		if (color) scheme[key] = color;
	}
	for (const diagram of diagramsIn(blocks)) {
		const name = diagram.parts.data?.partName;
		const original = name ? base.get(name) : undefined;
		if (
			!name ||
			done.has(name) ||
			(!diagram.inserted && (!original || semantic(diagram) === semantic(original)))
		)
			continue;
		done.add(name);
		const dataXml = await zip.file(name)?.async('string');
		if (!dataXml) continue;
		const read = async (key: 'layout' | 'colors' | 'quickStyle') => {
			const part = diagram.parts[key]?.partName;
			return part ? zip.file(part)?.async('string') : undefined;
		};
		const layoutXml = await read('layout');
		const colorsXml = await read('colors');
		const quickStyleXml = await read('quickStyle');
		const source = {
			...(layoutXml ? { layoutXml } : {}),
			...(colorsXml ? { colorsXml } : {}),
			...(quickStyleXml ? { quickStyleXml } : {}),
		};
		const cache = buildGeneratedDiagramCache(dataXml, {
			extentEmu: diagram.extentEmu,
			source,
			scheme,
			layoutId: diagram.layout?.uniqueId,
			family: diagram.layout?.family,
			minorFont: theme?.fonts.minor.latin,
			formatScheme: theme?.formatScheme,
		});
		if (!cache) continue;
		let index = 1;
		while (zip.file(`word/diagrams/drawing${index}.xml`)) index++;
		const drawingName = `word/diagrams/drawing${index}.xml`;
		const relationship = await ensureDocumentRelationship(
			zip,
			DIAGRAM_RELATIONSHIP_TYPES.drawing,
			drawingName.replace(/^word\//, ''),
		);
		zip.file(name, withDiagramDrawingRelationship(cache.dataXml, relationship));
		zip.file(drawingName, cache.drawingXml);
		await ensureContentTypeOverride(zip, drawingName, DIAGRAM_CONTENT_TYPES.drawing);
	}
}
