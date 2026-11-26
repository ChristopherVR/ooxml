import { relativePartTarget } from '../../opc/clone-part-graph';
import { RELATIONSHIP_TYPES } from '../../opc/relationship-types';
import { NS } from '../../xml/namespaces';
import { buildXml, children, first, parseXml, type XmlElement } from '../../xml/xml';
import type { SourceIndex } from '../read/package';
import { richPartName } from '../read/rich-values';
import { LOCAL_IMAGE_KEY, RICH_PARTS, type EmbeddedCellPicture } from '../rich-values';
import type { MetadataPlan } from './metadata';
import { type PackageWriter, RelationshipSet } from './package-writer';
import { encodeEscapes } from './xml-out';

type Part = (typeof RICH_PARTS)[keyof typeof RICH_PARTS];
type WritePart = (part: Part, existing: string | undefined, root: XmlElement) => string;
type NewValue =
	| { readonly kind: 'picture'; readonly picture: EmbeddedCellPicture }
	| { readonly kind: 'error'; readonly id: string };

function append(root: XmlElement, local: string, before: XmlElement | null = null): XmlElement {
	const node = root.ownerDocument.createElementNS(root.namespaceURI, local);
	root.insertBefore(node, before);
	return node;
}

/** Appends new rich values, keeping every pre-existing rich-value and metadata index. */
export class RichValueWritePlan {
	private readonly pending = new Map<string, { value: NewValue; vm: number }>();
	private readonly dataPart: string | undefined;
	private data: XmlElement | undefined;
	private recordOffset = 0;

	constructor(
		private readonly writer: PackageWriter,
		private readonly source: SourceIndex | undefined,
		private readonly book: string,
		private readonly metadataPart: string | undefined,
		private readonly metadata: MetadataPlan,
		private readonly reserved: ReadonlySet<string>,
	) {
		this.dataPart = this.name(RICH_PARTS.data);
	}

	private name(part: Part): string | undefined {
		return richPartName(this.source, this.book, this.metadataPart, part);
	}

	private root(part: Part, name: string | undefined): XmlElement {
		const xml = name && this.source?.text(name);
		const root = parseXml(xml || `<${part.root} xmlns="${part.ns}"/>`, {
			label: 'XLSX rich data',
			limits: { maxNodes: 250_000, maxDepth: 64 },
		}).documentElement;
		if (root.localName !== part.root || root.namespaceURI !== part.ns)
			throw new Error('Cannot add a picture to a malformed rich-data part.');
		return root;
	}

	add(value: NewValue): number {
		const key = JSON.stringify(value);
		const existing = this.pending.get(key);
		if (existing) return existing.vm;
		if (!this.data) {
			this.data = this.root(RICH_PARTS.data, this.dataPart);
			this.recordOffset = children(this.data, 'rv', RICH_PARTS.data.ns).length;
		}
		const vm = this.metadata.addRichValue(this.recordOffset + this.pending.size);
		this.pending.set(key, { value, vm });
		return vm;
	}

	write(bookRels: RelationshipSet): void {
		if (!this.data) return;
		const write: WritePart = (part, existing, root) => {
			const name =
				existing ??
				this.writer.uniqueName(
					(n) => `xl/richData/${n === 1 ? part.file : part.file.replace('.xml', `${n}.xml`)}`,
					this.reserved,
				);
			if (existing) this.writer.carry(existing);
			this.writer.add(name, buildXml(root.ownerDocument), part.type);
			if (!existing) bookRels.add(part.rel, relativePartTarget(this.book, name));
			return name;
		};
		const structureName = this.name(RICH_PARTS.structures);
		const structures = this.root(RICH_PARTS.structures, structureName);
		const entries = children(structures, 's', RICH_PARTS.structures.ns);
		const oldStructureCount = entries.length;
		const values = [...this.pending.values()].map(({ value }) => value);
		const pictures = values.flatMap((value) => (value.kind === 'picture' ? [value.picture] : []));
		const images = pictures.length
			? this.writeImageRelations(pictures, write)
			: new Map<string, number>();
		const imageStructure = pictures.length
			? ensureStructure(structures, entries, '_localImage', [
					[LOCAL_IMAGE_KEY, 'i'],
					['CalcOrigin', 'i'],
					['Text', 's'],
				])
			: -1;
		const errorStructure = values.some((value) => value.kind === 'error')
			? ensureStructure(structures, entries, '_error', [['errorType', 'i']])
			: -1;
		if (entries.length !== oldStructureCount) {
			structures.setAttribute('count', String(entries.length));
			write(RICH_PARTS.structures, structureName, structures);
		}
		const dataExt = first(this.data, 'extLst', RICH_PARTS.data.ns) ?? null;
		for (const value of values) {
			const record = append(this.data, 'rv', dataExt);
			record.setAttribute('s', String(value.kind === 'picture' ? imageStructure : errorStructure));
			const fields =
				value.kind === 'picture'
					? [String(images.get(value.picture.partName)), '5', encodeEscapes(value.picture.altText)]
					: [value.id];
			for (const field of fields) append(record, 'v').textContent = field;
		}
		this.data.setAttribute('count', String(this.recordOffset + values.length));
		write(RICH_PARTS.data, this.dataPart, this.data);
		if (!this.name(RICH_PARTS.types))
			write(RICH_PARTS.types, undefined, this.root(RICH_PARTS.types, undefined));
	}

	private writeImageRelations(
		pictures: readonly EmbeddedCellPicture[],
		write: WritePart,
	): Map<string, number> {
		const relPart = RICH_PARTS.relationships;
		const oldRelName = this.name(relPart);
		const relRoot = this.root(relPart, oldRelName);
		const relName =
			oldRelName ??
			this.writer.uniqueName(
				(n) => `xl/richData/${n === 1 ? relPart.file : `richValueRel${n}.xml`}`,
				this.reserved,
			);
		const rels = new RelationshipSet();
		if (oldRelName && this.source)
			for (const [id, rel] of this.source.rels(oldRelName)) rels.keep(id, rel);
		const images = new Map<string, number>();
		const existingRels = children(relRoot, 'rel', relPart.ns);
		existingRels.forEach((rel, index) => {
			const relation = oldRelName
				? this.source?.rels(oldRelName).get(rel.getAttributeNS(NS.r, 'id') ?? '')
				: undefined;
			if (oldRelName && relation?.type === RELATIONSHIP_TYPES.image) {
				const target = this.source?.target(oldRelName, relation);
				if (target) images.set(target, index);
			}
		});
		let relationIndex = existingRels.length;
		const relExt = first(relRoot, 'extLst', relPart.ns) ?? null;
		for (const { partName } of pictures) {
			if (images.has(partName)) continue;
			images.set(partName, relationIndex++);
			this.writer.carry(partName);
			const id = rels.add(RELATIONSHIP_TYPES.image, relativePartTarget(relName, partName));
			append(relRoot, 'rel', relExt).setAttributeNS(NS.r, 'r:id', id);
		}
		if (relationIndex !== existingRels.length) {
			write(relPart, oldRelName, relRoot);
			this.writer.rels(relName, rels);
		}
		return images;
	}
}

function ensureStructure(
	root: XmlElement,
	entries: XmlElement[],
	type: string,
	keys: ReadonlyArray<readonly [string, string]>,
): number {
	const existing = entries.findIndex((entry) => {
		if (entry.getAttribute('t') !== type) return false;
		const fields = children(entry, 'k', RICH_PARTS.structures.ns);
		return (
			fields.length === keys.length &&
			keys.every(
				([name, kind], i) =>
					fields[i]?.getAttribute('n') === name && fields[i]?.getAttribute('t') === kind,
			)
		);
	});
	if (existing >= 0) return existing;
	const entry = append(root, 's', first(root, 'extLst', RICH_PARTS.structures.ns) ?? null);
	entry.setAttribute('t', type);
	for (const [name, kind] of keys) {
		const key = append(entry, 'k');
		key.setAttribute('n', name);
		key.setAttribute('t', kind);
	}
	return entries.push(entry) - 1;
}
