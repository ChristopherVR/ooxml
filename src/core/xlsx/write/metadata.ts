// `xl/metadata.xml` on save: the source part is kept (so `vm` value metadata and other `cm`
// cell metadata indices stay valid) and the dynamic-array (XLDAPR) block is merged into it.
import { NS, buildXml, elements, parseXml, type XmlElement } from '../../xml/index';
import { RELATIONSHIP_TYPES } from '../../opc/index';
import { DYNAMIC_ARRAY_TYPE, parseDynamicArrayMetadata } from '../read/metadata';
import { CONTENT_TYPES } from '../read/package';
import { RICH_METADATA, RICH_METADATA_EXT, RICH_NS } from '../rich-values';
import { relativeTarget } from './drawing';
import { DYNAMIC_ARRAY_METADATA_XML } from './dynamic-array';
import type { PackageWriter, RelationshipSet } from './package-writer';

const NS_XDA = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const XLDAPR_EXT = '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}';
/** CT_Metadata child order. */
const ORDER = [
	'metadataTypes',
	'metadataStrings',
	'mdxMetadata',
	'futureMetadata',
	'cellMetadata',
	'valueMetadata',
	'extLst',
];

const kids = (parent: XmlElement, local: string) =>
	elements(parent).filter((e) => e.localName === local);

function child(doc: Document, root: XmlElement, local: string): XmlElement {
	const existing = kids(root, local)[0];
	if (existing) return existing;
	const node = doc.createElementNS(NS.x, local);
	const rank = ORDER.indexOf(local);
	const next = elements(root).find((e) => ORDER.indexOf(e.localName) > rank);
	root.insertBefore(node, next ?? null);
	return node;
}

const setCount = (node: XmlElement, local: string) =>
	node.setAttribute('count', String(kids(node, local).length));

function appendRichMetadata(xml: string, indices: readonly number[]): string {
	const doc = parseXml(xml, { label: 'XLSX metadata' });
	const root = doc.documentElement;
	const types = child(doc, root, 'metadataTypes');
	const typeList = kids(types, 'metadataType');
	let typeIndex = typeList.findIndex((node) => node.getAttribute('name') === RICH_METADATA);
	if (typeIndex < 0) {
		typeIndex = typeList.length;
		const type = doc.createElementNS(NS.x, 'metadataType');
		type.setAttribute('name', RICH_METADATA);
		type.setAttribute('minSupportedVersion', '120000');
		for (const flag of 'copy pasteAll pasteValues merge splitFirst rowColShift clearFormats clearComments assign coerce'.split(
			' ',
		))
			type.setAttribute(flag, '1');
		types.appendChild(type);
	}
	setCount(types, 'metadataType');
	let future = kids(root, 'futureMetadata').find(
		(node) => node.getAttribute('name') === RICH_METADATA,
	);
	if (!future) {
		future = doc.createElementNS(NS.x, 'futureMetadata');
		future.setAttribute('name', RICH_METADATA);
		const next = elements(root).find(
			(node) => ORDER.indexOf(node.localName) > ORDER.indexOf('futureMetadata'),
		);
		root.insertBefore(future, next ?? null);
	}
	const values = child(doc, root, 'valueMetadata');
	const offset = kids(future, 'bk').length;
	indices.forEach((richIndex, index) => {
		const bk = doc.createElementNS(NS.x, 'bk');
		const extList = doc.createElementNS(NS.x, 'extLst');
		const ext = doc.createElementNS(NS.x, 'ext');
		ext.setAttribute('uri', RICH_METADATA_EXT);
		const rvb = doc.createElementNS(RICH_NS, 'xlrd:rvb');
		rvb.setAttribute('i', String(richIndex));
		ext.appendChild(rvb);
		extList.appendChild(ext);
		bk.appendChild(extList);
		future.appendChild(bk);
		const value = doc.createElementNS(NS.x, 'bk');
		const rc = doc.createElementNS(NS.x, 'rc');
		rc.setAttribute('t', String(typeIndex + 1));
		rc.setAttribute('v', String(offset + index));
		value.appendChild(rc);
		values.appendChild(value);
	});
	setCount(future, 'bk');
	setCount(values, 'bk');
	return buildXml(doc);
}

/**
 * Adds an XLDAPR metadata type, future-metadata block and cell-metadata block to a source
 * `metadata.xml`; returns the merged XML and the new one-based `cm` index, or `undefined` when
 * the source cannot be read.
 */
function mergeDynamicArray(xml: string): { xml: string; cm: number } | undefined {
	let doc: Document;
	try {
		doc = parseXml(xml, { label: 'XLSX metadata' });
	} catch {
		return undefined;
	}
	const root = doc.documentElement;
	if (root.localName !== 'metadata') return undefined;
	const types = child(doc, root, 'metadataTypes');
	let typeIndex = kids(types, 'metadataType').findIndex(
		(t) => t.getAttribute('name') === DYNAMIC_ARRAY_TYPE,
	);
	if (typeIndex < 0) {
		const type = doc.createElementNS(NS.x, 'metadataType');
		const flags =
			'copy pasteAll pasteValues merge splitFirst rowColShift clearFormats clearComments assign coerce cellMeta';
		type.setAttribute('name', DYNAMIC_ARRAY_TYPE);
		type.setAttribute('minSupportedVersion', '120000');
		for (const flag of flags.split(' ')) type.setAttribute(flag, '1');
		types.appendChild(type);
		typeIndex = kids(types, 'metadataType').length - 1;
	}
	setCount(types, 'metadataType');
	let future = kids(root, 'futureMetadata').find(
		(f) => f.getAttribute('name') === DYNAMIC_ARRAY_TYPE,
	);
	if (!future) {
		future = doc.createElementNS(NS.x, 'futureMetadata');
		future.setAttribute('name', DYNAMIC_ARRAY_TYPE);
		const after = elements(root).find(
			(e) => ORDER.indexOf(e.localName) > ORDER.indexOf('futureMetadata'),
		);
		root.insertBefore(future, after ?? null);
	}
	if (root.lookupNamespaceURI('xda') !== NS_XDA)
		root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:xda', NS_XDA);
	const bk = doc.createElementNS(NS.x, 'bk');
	const extLst = doc.createElementNS(NS.x, 'extLst');
	const ext = doc.createElementNS(NS.x, 'ext');
	ext.setAttribute('uri', XLDAPR_EXT);
	const props = doc.createElementNS(NS_XDA, 'xda:dynamicArrayProperties');
	props.setAttribute('fDynamic', '1');
	props.setAttribute('fCollapsed', '0');
	ext.appendChild(props);
	extLst.appendChild(ext);
	bk.appendChild(extLst);
	future.appendChild(bk);
	setCount(future, 'bk');
	const cells = child(doc, root, 'cellMetadata');
	const cellBk = doc.createElementNS(NS.x, 'bk');
	const rc = doc.createElementNS(NS.x, 'rc');
	rc.setAttribute('t', String(typeIndex + 1));
	rc.setAttribute('v', String(kids(future, 'bk').length - 1));
	cellBk.appendChild(rc);
	cells.appendChild(cellBk);
	setCount(cells, 'bk');
	return { xml: buildXml(doc), cm: kids(cells, 'bk').length };
}

/**
 * What the writer does with cell metadata. With a source part, its indices stay valid: `vm`
 * and non-dynamic `cm` values are written back, and the dynamic-array block is the source's
 * own XLDAPR block or one appended to it.
 */
export class MetadataPlan {
	/** The one-based `cm` index that marks a dynamic-array formula. */
	readonly dynamicCm: number;
	/** Whether source `vm`/`cm` indices may be written (the source part is kept). */
	readonly keepsSource: boolean;
	private readonly merged: string | undefined;
	private readonly richIndices: number[] = [];
	private sourceValueCount: number | undefined;

	constructor(private readonly sourceXml: string | undefined) {
		const existing = [...parseDynamicArrayMetadata(sourceXml)][0];
		const merged = sourceXml && existing === undefined ? mergeDynamicArray(sourceXml) : undefined;
		this.keepsSource = sourceXml !== undefined && (existing !== undefined || merged !== undefined);
		this.dynamicCm = existing ?? merged?.cm ?? 1;
		this.merged = merged?.xml;
	}

	/** The metadata part to write, or `undefined` when none is needed. */
	xml(dynamicArraysUsed: boolean): string | undefined {
		const xml = this.keepsSource
			? dynamicArraysUsed
				? (this.merged ?? this.sourceXml)
				: this.sourceXml
			: dynamicArraysUsed
				? DYNAMIC_ARRAY_METADATA_XML
				: undefined;
		return this.richIndices.length
			? appendRichMetadata(xml ?? `<metadata xmlns="${NS.x}"/>`, this.richIndices)
			: xml;
	}

	addRichValue(index: number): number {
		if (this.sourceValueCount === undefined) {
			const root =
				this.keepsSource && this.sourceXml ? parseXml(this.sourceXml).documentElement : undefined;
			const values = root && kids(root, 'valueMetadata')[0];
			this.sourceValueCount = values ? kids(values, 'bk').length : 0;
		}
		return this.sourceValueCount + this.richIndices.push(index);
	}
}

/**
 * Writes the workbook's `metadata.xml` (if one is needed) at the source part's name, or a new
 * one, and adds its relationship. `sourcePart` is the source workbook's sheetMetadata target.
 */
export function writeMetadataPart(
	writer: PackageWriter,
	rels: RelationshipSet,
	bookPart: string,
	plan: MetadataPlan,
	dynamicArraysUsed: boolean,
	sourcePart: string | undefined,
	reserved: ReadonlySet<string>,
): void {
	const xml = plan.xml(dynamicArraysUsed);
	if (xml === undefined) return;
	const part =
		sourcePart && plan.keepsSource
			? sourcePart
			: writer.uniqueName(
					(n) => (n === 1 ? 'xl/metadata.xml' : `xl/metadata${n}.xml`),
					new Set([...reserved].filter((p) => p !== sourcePart)),
				);
	if (sourcePart === part) writer.carry(part);
	if (!writer.has(part)) writer.add(part, xml, CONTENT_TYPES.sheetMetadata);
	else if (xml !== writer.source?.text(part)) writer.replace(part, xml);
	rels.add(RELATIONSHIP_TYPES.sheetMetadata, relativeTarget(bookPart, part));
}
