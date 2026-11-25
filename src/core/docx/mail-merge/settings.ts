import type JSZip from 'jszip';
import {
	RELATIONSHIP_TYPES,
	nextRelationshipId,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
	type Relationship,
} from '../../opc/index';
import { NS, elements } from '../../xml/index';
import { orderChildren } from '../element-order';
import { DOCUMENT_PART, DOCUMENT_RELS, SETTINGS_PART } from '../package-information';
import { insertSetting, SETTINGS_CONTENT_TYPE } from '../settings';
import { parseUnsignedInteger } from '../simple-types';
import {
	buildXml,
	children,
	first,
	getR,
	getW,
	makeW,
	parseXml,
	WORD_NS,
	type XmlElement,
} from '../xml';
import { ensureContentTypeOverride, ensureDocumentRelationship } from '../zip-parts';

const SOURCE_TYPE = `${NS.r}/mailMergeSource`;
const VALUES = ['mainDocumentType', 'dataType', 'query'] as const;
const MAIL_MERGE_ORDER = [
	'mainDocumentType',
	'linkToQuery',
	'dataType',
	'connectString',
	'query',
	'dataSource',
	'headerSource',
	'doNotSuppressBlankLines',
	'destination',
	'addressFieldName',
	'mailSubject',
	'mailAsAttachment',
	'viewMergedData',
	'activeRecord',
	'checkErrors',
	'odso',
];
const ODSO_ORDER = [
	'udl',
	'table',
	'src',
	'colDelim',
	'type',
	'fHdr',
	'fieldMapData',
	'recipientData',
];
const MAP_ORDER = ['type', 'name', 'mappedName', 'column', 'lid', 'dynamicAddress'];

/** A link stored in the settings part. The merge engine treats its target as metadata. */
export interface MailMergeSourceReference {
	relationshipId?: string;
	target?: string;
	targetMode?: 'External' | 'Internal';
}

/** One ODSO entry, in source order. Unmapped entries can occupy a position in Word's field list. */
export interface MailMergeFieldMapping {
	type?: string;
	name?: string;
	mappedName?: string;
	/** Zero-based column position. */
	column?: number;
}

export interface MailMergeSettings {
	mainDocumentType?: string;
	dataType?: string;
	query?: string;
	dataSource?: MailMergeSourceReference;
	odsoSource?: MailMergeSourceReference;
	fieldMappings?: readonly MailMergeFieldMapping[];
}

function settingsPath(rels: ReadonlyMap<string, Relationship>): string {
	const rel = [...rels.values()].find(
		(item) => item.type === RELATIONSHIP_TYPES.settings && item.mode === 'Internal',
	);
	return (rel && resolvePartPath(DOCUMENT_PART, rel.target)) || SETTINGS_PART;
}

const value = (parent: XmlElement, name: string) => getW(first(parent, name), 'val');

export async function readMailMergeSettings(
	zip: JSZip,
	documentRelationships: ReadonlyMap<string, Relationship>,
	canonicalSettings: string | undefined,
): Promise<MailMergeSettings | undefined> {
	const path = settingsPath(documentRelationships);
	const xml = path === SETTINGS_PART ? canonicalSettings : await zip.file(path)?.async('string');
	const merge = xml && first(parseXml(xml).documentElement, 'mailMerge');
	if (!merge) return undefined;
	const settings: MailMergeSettings = {};
	for (const key of VALUES) {
		const text = value(merge, key);
		if (text !== undefined) settings[key] = text;
	}
	const odso = first(merge, 'odso');
	const rels = parseRelationships(await zip.file(relationshipsPartFor(path))?.async('string'));
	const ref = (node: XmlElement | undefined): MailMergeSourceReference | undefined => {
		const relationshipId = getR(node, 'id');
		if (!relationshipId) return undefined;
		const rel = rels.get(relationshipId);
		return rel ? { relationshipId, target: rel.target, targetMode: rel.mode } : { relationshipId };
	};
	const dataSource = ref(first(merge, 'dataSource'));
	const odsoSource = ref(first(odso, 'src'));
	if (dataSource) settings.dataSource = dataSource;
	if (odsoSource) settings.odsoSource = odsoSource;
	if (odso) {
		const maps = children(odso, 'fieldMapData');
		if (maps.length)
			settings.fieldMappings = maps.map((element) => {
				const mapping: MailMergeFieldMapping = {};
				for (const key of ['type', 'name', 'mappedName'] as const) {
					const text = value(element, key);
					if (text !== undefined) mapping[key] = text;
				}
				const column = parseUnsignedInteger(value(element, 'column'));
				if (column !== undefined) mapping.column = column;
				return mapping;
			});
	}
	return settings;
}

function setValue(parent: XmlElement, name: string, next: string | number | undefined): void {
	const previous = first(parent, name);
	if (next === undefined) {
		if (previous) parent.removeChild(previous);
		return;
	}
	const child = previous ?? makeW(parent.ownerDocument, name);
	child.setAttributeNS(WORD_NS, 'w:val', String(next));
	if (!previous) parent.appendChild(child);
}

function relationshipIds(element: XmlElement): Set<string> {
	return new Set(
		[element, ...Array.from(element.getElementsByTagName('*'))].flatMap((node) =>
			Array.from(node.attributes)
				.filter((attribute) => attribute.namespaceURI === NS.r)
				.map((attribute) => attribute.value),
		),
	);
}

function writeFieldMappings(
	odso: XmlElement,
	mappings: readonly MailMergeFieldMapping[],
	base: readonly MailMergeFieldMapping[],
): void {
	const previous = children(odso, 'fieldMapData');
	mappings.forEach((mapping, index) => {
		const existing = previous[index];
		const node = existing ?? makeW(odso.ownerDocument, 'fieldMapData');
		for (const key of ['type', 'name', 'mappedName', 'column'] as const)
			if (mapping[key] !== base[index]?.[key]) setValue(node, key, mapping[key]);
		orderChildren(node, MAP_ORDER);
		if (!existing) odso.appendChild(node);
	});
	for (const obsolete of previous.slice(mappings.length)) odso.removeChild(obsolete);
}

/** Patches modeled settings and leaves other settings, ODSO metadata and relationships in place. */
export async function writeMailMergeSettings(
	zip: JSZip,
	next: MailMergeSettings | undefined,
	base: MailMergeSettings | undefined,
): Promise<void> {
	if (JSON.stringify(next) === JSON.stringify(base)) return;
	const path = settingsPath(parseRelationships(await zip.file(DOCUMENT_RELS)?.async('string')));
	const file = zip.file(path);
	if (!file && !next) return;
	const doc = parseXml((await file?.async('string')) ?? `<w:settings xmlns:w="${WORD_NS}"/>`);
	const root = doc.documentElement;
	const original = first(root, 'mailMerge');
	const oldIds = original ? relationshipIds(original) : new Set<string>();
	const relsPath = relationshipsPartFor(path);
	const rels = parseXml(
		(await zip.file(relsPath)?.async('string')) ?? `<Relationships xmlns="${NS.rels}"/>`,
	);
	const byId = new Map(
		elements(rels.documentElement).flatMap((element) => {
			const id = element.getAttribute('Id');
			return element.namespaceURI === NS.rels && element.localName === 'Relationship' && id
				? [[id, element] as const]
				: [];
		}),
	);
	let relsChanged = false;
	const setReference = (
		parent: XmlElement,
		name: string,
		next: MailMergeSourceReference | undefined,
	) => {
		const existing = first(parent, name);
		if (!next) {
			if (existing) parent.removeChild(existing);
			return;
		}
		let id = next.relationshipId ?? getR(existing, 'id');
		if (next.target !== undefined) {
			let relation = id ? byId.get(id) : undefined;
			const mode =
				next.targetMode ??
				(relation ? (relation.getAttribute('TargetMode') ?? 'Internal') : 'External');
			if (relation && relation.getAttribute('Type') !== SOURCE_TYPE)
				throw new Error(`Relationship ${id} is already used by another settings feature.`);
			if (
				relation &&
				(relation.getAttribute('Target') !== next.target ||
					(relation.getAttribute('TargetMode') ?? 'Internal') !== mode)
			) {
				// The current id may still be used by another source or settings entry.
				relation = undefined;
				id = undefined;
			}
			if (!relation) {
				id ??= nextRelationshipId(new Set(byId.keys()));
				relation = rels.createElementNS(NS.rels, 'Relationship');
				relation.setAttribute('Id', id);
				relation.setAttribute('Type', SOURCE_TYPE);
				relation.setAttribute('Target', next.target);
				if (mode === 'External') relation.setAttribute('TargetMode', mode);
				rels.documentElement.appendChild(relation);
				byId.set(id, relation);
				relsChanged = true;
			}
		}
		if (!id) throw new Error('A mail merge source needs a relationship id or target.');
		const child = existing ?? makeW(doc, name);
		child.setAttributeNS(NS.r, 'r:id', id);
		if (!existing) parent.appendChild(child);
	};
	if (next) {
		const merge = original ?? makeW(doc, 'mailMerge');
		for (const key of VALUES)
			if (!original || next[key] !== base?.[key]) setValue(merge, key, next[key]);
		if (!original && !next.mainDocumentType) setValue(merge, 'mainDocumentType', 'formLetters');
		if (!original && !next.dataType) setValue(merge, 'dataType', 'textFile');
		if (JSON.stringify(next.dataSource) !== JSON.stringify(base?.dataSource))
			setReference(merge, 'dataSource', next.dataSource);
		const existingOdso = first(merge, 'odso');
		if (existingOdso || next.odsoSource || next.fieldMappings?.length) {
			const odso = existingOdso ?? makeW(doc, 'odso');
			if (JSON.stringify(next.odsoSource) !== JSON.stringify(base?.odsoSource))
				setReference(odso, 'src', next.odsoSource);
			if (JSON.stringify(next.fieldMappings) !== JSON.stringify(base?.fieldMappings))
				writeFieldMappings(odso, next.fieldMappings ?? [], base?.fieldMappings ?? []);
			orderChildren(odso, ODSO_ORDER);
			if (!existingOdso) merge.appendChild(odso);
		}
		orderChildren(merge, MAIL_MERGE_ORDER);
		if (!original) insertSetting(root, merge);
	} else if (original) root.removeChild(original);
	const remainingIds = relationshipIds(root);
	for (const id of oldIds) {
		const relation = byId.get(id);
		if (relation?.getAttribute('Type') === SOURCE_TYPE && !remainingIds.has(id)) {
			rels.documentElement.removeChild(relation);
			relsChanged = true;
		}
	}
	if (relsChanged) zip.file(relsPath, buildXml(rels));
	zip.file(path, buildXml(doc));
	if (!file) {
		await ensureContentTypeOverride(zip, path, SETTINGS_CONTENT_TYPE);
		if (path === SETTINGS_PART)
			await ensureDocumentRelationship(zip, RELATIONSHIP_TYPES.settings, 'settings.xml');
	}
}
