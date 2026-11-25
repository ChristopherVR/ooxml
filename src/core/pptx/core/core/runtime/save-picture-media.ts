import JSZip from 'jszip';
import { removeContentTypeOverride } from '../../../../opc/package';
import { partNameKey } from '../../../../opc/part-names';
import {
	buildRelationshipsXml,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
	type Relationship,
} from '../../../../opc/relationships';
import { elements, parseXml, type XmlElement } from '../../../../xml';
import type { PptxElement, PptxSlide } from '../../types';

function hasEditedPicture(list: PptxElement[], depth = 0): boolean {
	return list.some((element) => {
		if (element.type === 'image' || element.type === 'picture') return element.imageDataEdited;
		return element.type === 'group' && depth < 32 && hasEditedPicture(element.children, depth + 1);
	});
}

function attributeValues(root: XmlElement): Set<string> {
	const pending = [root];
	const ids = new Set<string>();
	while (pending.length) {
		const node = pending.pop()!;
		for (const { value } of Array.from(node.attributes)) ids.add(value);
		for (const child of elements(node)) pending.push(child);
	}
	return ids;
}

function targetPart(source: string, relationship: Relationship): string | undefined {
	const target = relationship.target.split('#')[0];
	return relationship.mode === 'Internal' && target ? resolvePartPath(source, target) : undefined;
}

/**
 * Remove detached raster parts from the export. Keep the editor's archive and source relationships
 * intact because its undo history can still refer to the original blips.
 */
export async function archiveWithEditedPictures(
	zip: JSZip,
	slides: readonly PptxSlide[],
): Promise<JSZip> {
	const candidates = new Map<string, string>();
	const replacements = new Map<string, Map<string, Relationship>>();
	for (const slide of slides) {
		if (!hasEditedPicture(slide.elements)) continue;
		const path = relationshipsPartFor(slide.id);
		const rels = parseRelationships(await zip.file(path)?.async('string'));
		if (!rels.size) continue;
		const markup = await zip.file(slide.id)?.async('string');
		if (!markup) continue;
		const values = attributeValues(parseXml(markup).documentElement);
		let changed = false;
		for (const [id, rel] of rels) {
			if (!rel.type.endsWith('/image') || values.has(id)) continue;
			const part = targetPart(slide.id, rel);
			if (!part || !/^ppt\/media\/[^/]+\.(?:png|jpe?g)$/iu.test(part)) continue;
			if (zip.file(relationshipsPartFor(part))) continue;
			candidates.set(partNameKey(part), part);
			rels.delete(id);
			changed = true;
		}
		if (changed) replacements.set(path, rels);
	}
	if (!replacements.size) return zip;
	// Every remaining relationship counts, even ones from masters or unknown part types.
	for (const part of zip.file(/\.rels$/iu)) {
		const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/iu.exec(part.name);
		if (!match && part.name !== '_rels/.rels') continue;
		const owner = match?.[1] ? match[1] + '/' + match[2] : (match?.[2] ?? '');
		const rels = replacements.get(part.name) ?? parseRelationships(await part.async('string'));
		for (const rel of rels.values()) {
			const target = targetPart(owner, rel);
			if (target) candidates.delete(partNameKey(target));
		}
		if (!candidates.size) break;
	}
	const output = new JSZip();
	output.files = Object.assign(Object.create(null), zip.files);
	for (const [path, rels] of replacements)
		output.file(path, buildRelationshipsXml(rels), { createFolders: false });
	for (const name of candidates.values()) {
		output.remove(name);
		await removeContentTypeOverride(output, name);
	}
	return output;
}
