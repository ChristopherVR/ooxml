import type JSZip from 'jszip';
import { ensureContentTypeOverride } from './package';
import { contentTypeForPart, parseContentTypes } from './content-types';
import {
	buildRelationshipsXml,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
	type Relationship,
} from './relationships';

/** Make a package-relative target from a declaring part to another part. */
export function relativePartTarget(basePart: string, targetPart: string): string {
	const base = basePart.split('/').slice(0, -1);
	const target = targetPart.split('/');
	while (base.length > 0 && base[0] === target[0]) {
		base.shift();
		target.shift();
	}
	return [...base.map(() => '..'), ...target].join('/');
}

/** The first free `<stem><n><extension>` name for a copy of `source`, which `reserved` then holds. */
function allocatePartName(reserved: Set<string>, source: string): string {
	const dot = source.lastIndexOf('.');
	const extension =
		dot > source.lastIndexOf('/') && dot < source.length - 1 ? source.slice(dot) : '';
	let stemEnd = source.length - extension.length;
	if (extension) {
		while (stemEnd > 0) {
			const char = source.charCodeAt(stemEnd - 1);
			if (char < 48 || char > 57) break;
			stemEnd--;
		}
	}
	const stem = source.slice(0, stemEnd);
	let index = 1;
	while (reserved.has(`${stem}${index}${extension}`)) index += 1;
	const path = `${stem}${index}${extension}`;
	reserved.add(path);
	return path;
}

export interface ClonePartGraphOptions {
	zip: JSZip;
	/** Separate source archive for cross-package imports. Defaults to zip. */
	sourceZip?: JSZip;
	/** Shared source-to-target map, optionally seeded with all imported roots. */
	copies?: Map<string, string>;
	sourcePart: string;
	targetPart: string;
	/** Return false for shared infrastructure or references to other document parts. */
	shouldClone: (relationship: Relationship, targetPart: string) => boolean;
}

/**
 * Copy a part's relationship graph, keeping relationship IDs and external targets.
 * The caller owns the root part's content. Internal owned targets receive new
 * names, bytes, relationship parts and content-type overrides. A per-copy map
 * preserves shared dependencies and cycles, including backlinks to the root.
 */
export async function clonePartGraph(options: ClonePartGraphOptions): Promise<Map<string, string>> {
	const { zip, sourcePart, targetPart, shouldClone } = options;
	const sourceZip = options.sourceZip ?? zip;
	const copies = options.copies ?? new Map<string, string>();
	copies.set(sourcePart, targetPart);
	const types = parseContentTypes(await sourceZip.file('[Content_Types].xml')?.async('string'));
	const reserved = new Set(Object.keys(zip.files));
	reserved.add(targetPart);
	for (const target of copies.values()) reserved.add(target);

	const allocate = (source: string): string => allocatePartName(reserved, source);

	const copyRelationships = async (source: string, destination: string): Promise<void> => {
		const contentType = contentTypeForPart(types, source);
		if (contentType) await ensureContentTypeOverride(zip, destination, contentType);
		const xml = await sourceZip.file(relationshipsPartFor(source))?.async('string');
		if (!xml) return;
		const relationships = parseRelationships(xml);
		for (const relationship of relationships.values()) {
			if (relationship.mode === 'External') continue;
			const originalTarget = resolvePartPath(source, relationship.target);
			let newTarget = copies.get(originalTarget);
			if (!newTarget && shouldClone(relationship, originalTarget)) {
				const file = sourceZip.file(originalTarget);
				if (!file) throw new Error(`Cannot clone missing package part: ${originalTarget}`);
				newTarget = allocate(originalTarget);
				copies.set(originalTarget, newTarget);
				zip.file(newTarget, await file.async('uint8array'));
				await copyRelationships(originalTarget, newTarget);
			}
			relationship.target = relativePartTarget(destination, newTarget ?? originalTarget);
		}
		zip.file(relationshipsPartFor(destination), buildRelationshipsXml(relationships));
	};

	await copyRelationships(sourcePart, targetPart);
	return copies;
}

/**
 * Copy one part under a fresh name, with its bytes, content type and the owned part graph below it
 * (see {@link clonePartGraph}). Returns every source-to-copy path, the part itself included.
 */
export async function clonePart(
	options: Omit<ClonePartGraphOptions, 'sourcePart' | 'targetPart' | 'sourceZip'> & {
		part: string;
	},
): Promise<Map<string, string>> {
	const { zip, part } = options;
	const file = zip.file(part);
	if (!file) throw new Error(`Cannot clone missing package part: ${part}`);
	const reserved = new Set(Object.keys(zip.files));
	for (const target of options.copies?.values() ?? []) reserved.add(target);
	const targetPart = allocatePartName(reserved, part);
	zip.file(targetPart, await file.async('uint8array'));
	return clonePartGraph({ ...options, sourcePart: part, targetPart });
}
