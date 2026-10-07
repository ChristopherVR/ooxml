import type JSZip from 'jszip';
import { ensureContentTypeOverride } from './package';
import { parseContentTypes } from './content-types';
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

export interface ClonePartGraphOptions {
	zip: JSZip;
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
	const copies = new Map<string, string>([[sourcePart, targetPart]]);
	const types = parseContentTypes(await zip.file('[Content_Types].xml')?.async('string'));
	const reserved = new Set(Object.keys(zip.files));
	reserved.add(targetPart);

	const allocate = (source: string): string => {
		const match = /^(?<stem>.*?)(?:\d+)?(?<extension>\.[^/.]+)$/u.exec(source);
		const stem = match?.groups?.stem ?? source;
		const extension = match?.groups?.extension ?? '';
		let index = 1;
		while (reserved.has(`${stem}${index}${extension}`)) index += 1;
		const path = `${stem}${index}${extension}`;
		reserved.add(path);
		return path;
	};

	const copyRelationships = async (source: string, destination: string): Promise<void> => {
		const contentType = types.overrides.get(`/${source}`);
		if (contentType) await ensureContentTypeOverride(zip, destination, contentType);
		const xml = await zip.file(relationshipsPartFor(source))?.async('string');
		if (!xml) return;
		const relationships = parseRelationships(xml);
		for (const relationship of relationships.values()) {
			if (relationship.mode === 'External') continue;
			const originalTarget = resolvePartPath(source, relationship.target);
			let newTarget = copies.get(originalTarget);
			if (!newTarget && shouldClone(relationship, originalTarget)) {
				const file = zip.file(originalTarget);
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
