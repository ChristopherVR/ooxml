import { RELATIONSHIP_TYPES } from '../../opc/relationship-types';
import { NS } from '../../xml/namespaces';
import { children, first, parseXml, type XmlElement } from '../../xml/xml';
import {
	LOCAL_IMAGE_KEY,
	RICH_METADATA,
	RICH_NS,
	RICH_PARTS,
	richIndex,
	type EmbeddedCellPicture,
	type RichValueRecord,
	type RichValueStore,
} from '../rich-values';
import type { SourceIndex } from './package';
import { att, xText } from './xml-util';

type RichPart = (typeof RICH_PARTS)[keyof typeof RICH_PARTS];

export function richPartName(
	source: SourceIndex | undefined,
	owner: string | undefined,
	metadata: string | undefined,
	part: RichPart,
): string | undefined {
	if (!source) return undefined;
	return (
		(owner && source.targetOfType(owner, part.rel)) ||
		(metadata && source.targetOfType(metadata, part.rel)) ||
		undefined
	);
}

function fieldsFromMetadata(root: XmlElement): Map<number, number> {
	const types = first(root, 'metadataTypes', NS.x);
	const future = children(root, 'futureMetadata', NS.x).find(
		(node) => att(node, 'name') === RICH_METADATA,
	);
	const values = first(root, 'valueMetadata', NS.x);
	const indices = new Map<number, number>();
	if (!types || !future || !values) return indices;
	const names = children(types, 'metadataType', NS.x).map((type) => att(type, 'name'));
	const blocks = children(future, 'bk', NS.x).map((block) => {
		const extensions = first(block, 'extLst', NS.x);
		if (!extensions) return undefined;
		for (const extension of children(extensions, 'ext', NS.x)) {
			const index = richIndex(att(first(extension, 'rvb', RICH_NS), 'i'));
			if (index !== undefined) return index;
		}
		return undefined;
	});
	children(values, 'bk', NS.x).forEach((block, index) => {
		for (const record of children(block, 'rc', NS.x)) {
			const type = richIndex(att(record, 't'));
			const value = richIndex(att(record, 'v'));
			if (type === undefined || names[type - 1] !== RICH_METADATA || value === undefined) continue;
			const rich = blocks[value];
			if (rich !== undefined) indices.set(index + 1, rich);
			break;
		}
	});
	return indices;
}

function hiddenFields(root: XmlElement | undefined, flagName: string): Map<string, Set<string>> {
	const hidden = new Map<string, Set<string>>();
	if (!root) return hidden;
	const ns = RICH_PARTS.types.ns;
	const global = first(root, 'global', ns);
	const types = first(root, 'types', ns);
	for (const type of [...(global ? [global] : []), ...(types ? children(types, 'type', ns) : [])]) {
		const flags = first(type, 'keyFlags', ns);
		if (!flags) continue;
		const names = new Set<string>();
		for (const key of children(flags, 'key', ns)) {
			const name = att(key, 'name');
			if (
				name &&
				children(key, 'flag', ns).some(
					(flag) =>
						att(flag, 'name') === flagName && ['0', 'false'].includes(att(flag, 'value') ?? ''),
				)
			)
				names.add(name);
		}
		hidden.set(att(type, 'name') ?? '', names);
	}
	return hidden;
}

/** Resolves indices without expanding rich-value graphs or trusting declared counts. */
export function readRichValues(
	source: SourceIndex,
	workbook: string,
	metadata: string | undefined,
	maxXmlElements: number,
	warn: (message: string) => void,
): RichValueStore | undefined {
	const partName = (part: RichPart) => richPartName(source, workbook, metadata, part);
	const data = partName(RICH_PARTS.data);
	if (!data || !metadata) return undefined;
	const read = (name: string | undefined, tag: string, ns: string): XmlElement | undefined => {
		const xml = name && source.text(name);
		if (!xml) return undefined;
		try {
			const root = parseXml(xml, {
				label: 'XLSX rich data',
				limits: { maxNodes: Math.min(maxXmlElements, 250_000), maxDepth: 64 },
			}).documentElement;
			if (root.localName === tag && root.namespaceURI === ns) return root;
		} catch {
			// Keep the original bytes available for save even when the view cannot use them.
		}
		warn('An Excel rich-data part could not be read. Its original data is kept on save.');
		return undefined;
	};
	const metadataRoot = read(metadata, 'metadata', NS.x);
	const dataRoot = read(data, 'rvData', RICH_NS);
	const structuresRoot = read(partName(RICH_PARTS.structures), 'rvStructures', RICH_NS);
	if (!metadataRoot || !dataRoot || !structuresRoot) return undefined;
	const relName = partName(RICH_PARTS.relationships);
	const relRoot = read(relName, 'richValueRels', RICH_PARTS.relationships.ns);
	const media =
		relName && relRoot
			? children(relRoot, 'rel', RICH_PARTS.relationships.ns).map((node) => {
					const rel = source.rels(relName).get(node.getAttributeNS(NS.r, 'id') ?? '');
					return rel?.type === RELATIONSHIP_TYPES.image ? source.target(relName, rel) : undefined;
				})
			: [];
	const structures = children(structuresRoot, 's', RICH_NS).map((node) => ({
		type: att(node, 't') ?? '',
		keys: children(node, 'k', RICH_NS).map((key) => ({
			name: att(key, 'n') ?? '',
			type: att(key, 't') ?? '',
		})),
	}));
	const records = children(dataRoot, 'rv', RICH_NS).map((node): RichValueRecord | undefined => {
		const structureIndex = richIndex(att(node, 's'));
		const structure = structureIndex === undefined ? undefined : structures[structureIndex];
		if (!structure) return undefined;
		const fields = children(node, 'v', RICH_NS).flatMap((value, index) => {
			const key = structure.keys[index];
			return key ? [{ ...key, value: xText(value) }] : [];
		});
		let picture: EmbeddedCellPicture | undefined;
		if (structure.type === '_localImage') {
			const index = richIndex(
				fields.find((f) => f.name === LOCAL_IMAGE_KEY && f.type === 'i')?.value,
			);
			const partName = index === undefined ? undefined : media[index];
			if (partName)
				picture = {
					partName,
					altText: fields.find((f) => f.name === 'Text' && f.type === 's')?.value ?? '',
				};
		}
		return { type: structure.type, fields, ...(picture ? { picture } : {}) };
	});
	const types = read(partName(RICH_PARTS.types), 'rvTypesInfo', RICH_PARTS.types.ns);
	return {
		records,
		byMetadata: fieldsFromMetadata(metadataRoot),
		hiddenCardFields: hiddenFields(types, 'ShowInCardView'),
		hiddenFormulaFields: hiddenFields(types, 'ShowInDotNotation'),
	};
}
