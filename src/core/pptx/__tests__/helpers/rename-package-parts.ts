import JSZip from 'jszip';
import { buildXml, NS, parseXml } from '../../../xml/index';
import {
	buildRelationshipsXml,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
} from '../../../opc/relationships';
import { relativePartTarget } from '../../../opc/clone-part-graph';

/** Rename fixture parts and rebase every incoming and outgoing relationship. */
export async function renameFixtureParts(
	bytes: Uint8Array,
	names: ReadonlyMap<string, string>,
): Promise<Uint8Array> {
	const source = await JSZip.loadAsync(bytes),
		target = new JSZip();
	for (const [path, part] of Object.entries(source.files)) {
		if (part.dir) continue;
		if (path.endsWith('.rels')) {
			const owner =
				path === '_rels/.rels' ? '' : path.replace(/^(.*)\/_rels\/([^/]+)\.rels$/u, '$1/$2');
			const newOwner = names.get(owner) ?? owner;
			const relationships = parseRelationships(await part.async('string'));
			for (const relationship of relationships.values()) {
				if (relationship.mode === 'External') continue;
				const oldTarget = resolvePartPath(owner, relationship.target);
				relationship.target = relativePartTarget(newOwner, names.get(oldTarget) ?? oldTarget);
			}
			target.file(relationshipsPartFor(newOwner), buildRelationshipsXml(relationships));
		} else if (path === '[Content_Types].xml') {
			const doc = parseXml(await part.async('string'));
			for (const entry of Array.from(doc.getElementsByTagNameNS(NS.contentTypes, 'Override'))) {
				const replacement = names.get((entry.getAttribute('PartName') ?? '').slice(1));
				if (replacement) entry.setAttribute('PartName', `/${replacement}`);
			}
			target.file(path, buildXml(doc));
		} else target.file(names.get(path) ?? path, await part.async('uint8array'));
	}
	return target.generateAsync({ type: 'uint8array' });
}
