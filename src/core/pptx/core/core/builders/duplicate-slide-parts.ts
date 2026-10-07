import type JSZip from 'jszip';
import { clonePartGraph } from '../../../../opc/clone-part-graph';
import { cloneSlide } from '../../utils/clone-utils';
import type { PptxSlide } from '../../types';

// These targets belong to the presentation, not to the duplicated slide.
const SHARED_RELATIONSHIPS = new Set([
	'slide',
	'slideLayout',
	'slideMaster',
	'notesMaster',
	'handoutMaster',
	'theme',
	'hyperlink',
]);
const RAW_XML_KEYS = new Set(['rawXml', 'rawTiming', 'inkPartRawXml', 'rawMediaReferenceXml']);

/** Remap typed package paths without changing text, URLs or XML relationship IDs. */
export function remapPartPaths(value: unknown, copies: ReadonlyMap<string, string>): void {
	if (!value || typeof value !== 'object') return;
	for (const [key, child] of Object.entries(value)) {
		if (RAW_XML_KEYS.has(key)) continue;
		if (typeof child === 'string' && /path$/iu.test(key)) {
			const destination = copies.get(child);
			if (destination) (value as Record<string, unknown>)[key] = destination;
		} else {
			remapPartPaths(child, copies);
		}
	}
}

/** Copy owned parts, including chart dependencies and notes with their slide backlink. */
export async function duplicateSlideParts(
	zip: JSZip,
	slide: PptxSlide,
	sourcePart: string,
): Promise<Map<string, string>> {
	const copies = await clonePartGraph({
		zip,
		sourcePart,
		targetPart: slide.id,
		shouldClone: (relationship) =>
			!SHARED_RELATIONSHIPS.has(relationship.type.slice(relationship.type.lastIndexOf('/') + 1)),
	});
	// Some editor callers used shallow copies. Detach the model before changing
	// paths so a clone's save cannot redirect its source chart to the copied part.
	Object.assign(slide, cloneSlide(slide));
	remapPartPaths(slide, copies);
	return copies;
}
