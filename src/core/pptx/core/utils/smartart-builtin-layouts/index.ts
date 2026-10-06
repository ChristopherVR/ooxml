/**
 * PowerPoint's built-in SmartArt layout definitions, so a layout swap runs the real DiagramML
 * engine instead of a coarse family approximation.
 *
 * Switching layout used to clear the parsed `dgm:layoutDef` and fall back to one simplified
 * renderer per category (every "relationship" layout became the same radial drawing; timeline and
 * bending had their own simplified ones). Applying the target layout's own definition lets the
 * per-point engine (`smartart-engine/`), which matches PowerPoint's cached drawings, lay it out.
 *
 * The 176 definitions are about 350 KB gzipped, stored as per-layout gzip members in `data.ts` and
 * inflated on demand, so importing this module costs the catalogue only.
 *
 * @module utils/smartart-builtin-layouts
 */
import { XMLParser } from 'fast-xml-parser';

import type { PptxSmartArtData, XmlObject } from '../../types';
import { resolveSmartArtLayoutCategory } from '../../core/runtime/smartart-layout-category';
import { parseSmartArtLayoutDefinition } from '../smartart-layout-definition';
import { BUILTIN_SMARTART_LAYOUTS } from './catalog';
import { BUILTIN_SMARTART_LAYOUT_DATA } from './data';
import type { BuiltinSmartArtLayoutEntry } from './types';

export type { BuiltinSmartArtLayoutEntry } from './types';

/** The catalogue of built-in layouts (index only: no definition is inflated). */
export function listBuiltinSmartArtLayouts(): readonly BuiltinSmartArtLayoutEntry[] {
	return BUILTIN_SMARTART_LAYOUTS;
}

/** The catalogue entry for a layout `uniqueId`, or `undefined` when it is not built in. */
export function findBuiltinSmartArtLayout(id: string): BuiltinSmartArtLayoutEntry | undefined {
	return BUILTIN_SMARTART_LAYOUTS.find((entry) => entry.id === id);
}

let blob: Uint8Array | undefined;
function dataBlob(): Uint8Array {
	blob ??= Uint8Array.from(atob(BUILTIN_SMARTART_LAYOUT_DATA), (ch) => ch.charCodeAt(0));
	return blob;
}

async function gunzip(bytes: Uint8Array): Promise<string> {
	const stream = new Blob([bytes as BlobPart])
		.stream()
		.pipeThrough(new DecompressionStream('gzip'));
	return new Response(stream).text();
}

const inflated = new Map<string, Promise<string>>();

/** The `dgm:layoutDef` XML of a built-in layout; rejects for an unknown id. */
export function loadBuiltinSmartArtLayoutXml(id: string): Promise<string> {
	const entry = findBuiltinSmartArtLayout(id);
	if (!entry) {
		return Promise.reject(new Error(`Unknown built-in SmartArt layout: ${id}`));
	}
	let xml = inflated.get(id);
	if (!xml) {
		xml = gunzip(dataBlob().subarray(entry.offset, entry.offset + entry.length));
		inflated.set(id, xml);
	}
	return xml;
}

const localName = (key: string): string => key.slice(key.indexOf(':') + 1);

/** Parse a `dgm:layoutDef` XML into the typed definition the engine runs (raw XML attached). */
export function parseBuiltinLayoutDefinition(xml: string): PptxSmartArtData['layoutDefinition'] {
	const parser = new XMLParser({
		ignoreAttributes: false,
		attributeNamePrefix: '@_',
		parseAttributeValue: false,
		parseTagValue: false,
	});
	const root = Object.entries(parser.parse(xml) as XmlObject).find(
		([key]) => localName(key) === 'layoutDef',
	)?.[1] as XmlObject | undefined;
	const definition = parseSmartArtLayoutDefinition(root, localName);
	if (definition) {
		definition.rawXmlText = xml;
	}
	return definition;
}

/**
 * `data` re-laid-out as the built-in layout `id`: its definition replaces the current one and its
 * category becomes `resolvedLayoutType`. Nodes, connections, colours and styles are kept, and the
 * cached drawing is marked stale so it is rebuilt by the engine. Rejects for an unknown id.
 */
export async function applyBuiltinSmartArtLayout(
	data: PptxSmartArtData,
	id: string,
): Promise<PptxSmartArtData> {
	const definition = parseBuiltinLayoutDefinition(await loadBuiltinSmartArtLayoutXml(id));
	if (!definition) {
		throw new Error(`Built-in SmartArt layout has no layout node: ${id}`);
	}
	const resolved = resolveSmartArtLayoutCategory(
		id,
		(definition.categories ?? []).map((category) => category.type),
	);
	return {
		...data,
		layoutType: resolved ?? data.layoutType,
		resolvedLayoutType: resolved ?? data.resolvedLayoutType,
		layout: undefined,
		layoutDefinition: definition,
		// The definition came from the library, not the layout part the file carries.
		presLayoutVars: undefined,
		layoutDirty: true,
		drawingDirty: true,
		// Stale shapes are cleared (undefined stays undefined) so the engine rebuilds them.
		drawingShapes: data.drawingShapes && data.drawingShapes.length > 0 ? [] : undefined,
	};
}
