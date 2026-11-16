// The built-in SmartArt colour (`dgm:colorsDef`) and quick-style (`dgm:styleDef`) definitions,
// shared by every product: the full Change Colors and SmartArt Styles catalogue, inflated on
// demand. Microsoft's definitions, as PowerPoint writes them into a document that uses them (see
// PROVENANCE.md, "Built-in SmartArt colour and quick-style catalogue"). The gallery metadata is
// in the main `diagram` entry (`DIAGRAM_NATIVE_COLOR_ENTRIES`, `DIAGRAM_NATIVE_STYLE_ENTRIES`).
import { gunzipSync, strFromU8 } from 'fflate';
import type { DiagramNativeFormatChange } from '../edit/format';
import { BUILTIN_SMARTART_STYLE_DEFINITIONS } from './style-catalog';
import { BUILTIN_SMARTART_STYLE_DATA } from './style-data';
import type { BuiltinSmartArtDefinitionEntry } from './types';

export { BUILTIN_SMARTART_STYLE_DEFINITIONS } from './style-catalog';
export type { BuiltinSmartArtDefinitionEntry } from './types';

/**
 * The entry of a definition, by `uniqueId` or by its short id (`accent1_2`, `3d1`; the ids are
 * not all under one URN prefix: `3d8` is a 2009 definition).
 */
export function findBuiltinSmartArtStyleDefinition(
	kind: 'colors' | 'style',
	id: string,
): BuiltinSmartArtDefinitionEntry | undefined {
	return BUILTIN_SMARTART_STYLE_DEFINITIONS.find(
		(entry) =>
			entry.kind === kind &&
			(id.startsWith('urn:') ? entry.id === id : entry.id.split('/').at(-1) === id),
	);
}

let blob: Uint8Array | undefined;
const inflated = new Map<string, string>();

/** The definition XML (no XML declaration); throws for an id that is not built in. */
export function loadBuiltinSmartArtStyleDefinitionXml(
	kind: 'colors' | 'style',
	id: string,
): string {
	const entry = findBuiltinSmartArtStyleDefinition(kind, id);
	if (!entry) throw new Error(`Unknown built-in SmartArt ${kind} definition: ${id}`);
	let xml = inflated.get(entry.id);
	if (xml === undefined) {
		blob ??= Uint8Array.from(atob(BUILTIN_SMARTART_STYLE_DATA), (ch) => ch.charCodeAt(0));
		xml = strFromU8(gunzipSync(blob.subarray(entry.offset, entry.offset + entry.length)));
		inflated.set(entry.id, xml);
	}
	return xml;
}

/** The format change that applies a built-in definition (Change Colors or SmartArt Styles). */
export function builtinSmartArtFormatChange(
	kind: 'colors' | 'style',
	id: string,
): DiagramNativeFormatChange {
	const definitionXml = loadBuiltinSmartArtStyleDefinitionXml(kind, id);
	return { kind, value: id.split('/').at(-1) ?? id, definitionXml };
}
