// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Assigns fresh relationship ids when write.ts introduces a new external hyperlink or a
// newly inserted picture, and records what save.ts must add to word/_rels/document.xml.rels.
import { RELATIONSHIP_TYPES, isSafeHyperlinkHref, nextRelationshipId } from '../opc/index';
import type { ChartExSpace, ChartSpace } from '../chart/index';
import { INK_RELATIONSHIP, type DocxInk } from './ink';
import { DocPrIdAllocator } from './docpr-ids';
import { isWordNavigableHref } from './hyperlink';
import { isElement, REL_NS, type XmlDocument } from './xml';

export const HYPERLINK_RELATIONSHIP_TYPE = RELATIONSHIP_TYPES.hyperlink;
export const IMAGE_RELATIONSHIP_TYPE = RELATIONSHIP_TYPES.image;
export const CHART_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
export const CHART_EX_RELATIONSHIP_TYPE =
	'http://schemas.microsoft.com/office/2014/relationships/chartEx';
export { isSafeHyperlinkHref };

export interface NewRelationship {
	id: string;
	type: string;
	target: string;
	mode: 'Internal' | 'External';
	/** Absolute package part name (e.g. `word/media/image1.png`), set only for newly added media parts. */
	partName?: string;
	/** A new XML part: its content type goes into an `Override`, not an extension `Default`. */
	overrideContentType?: string;
	/** The new part's text, when the writer produced it (otherwise it comes from pending media). */
	content?: string;
	/** A chart inserted in this session: its part, relationships and workbook are created on save, which sets the target. */
	chart?: ChartSpace;
	/** An extended chart inserted in this session, written the same way as {@link chart}. */
	chartEx?: ChartExSpace;
	/** Ink drawn in this session: its InkML part (`word/ink/inkN.xml`) is written and named on save. */
	ink?: DocxInk;
}

/** A part a new diagram adds: its relationship type and, for the data part, its text. */
export interface NewDiagramPart {
	partName: string;
	type: string;
	contentType: string;
	content?: string;
}

/** Walks every element/attribute in a parsed document.xml to find relationship ids already in use. */
export function scanUsedRelationshipIds(doc: XmlDocument): Set<string> {
	const used = new Set<string>();
	// Iterative and allocation-free: this runs over the whole body on every save.
	const stack: Node[] = doc.documentElement ? [doc.documentElement] : [];
	while (stack.length) {
		const node = stack.pop()!;
		if (!isElement(node)) continue;
		const attributes = node.attributes;
		for (let index = 0; index < attributes.length; index++) {
			const attribute = attributes.item(index);
			if (attribute?.namespaceURI === REL_NS) used.add(attribute.value);
		}
		for (let child = node.firstChild; child; child = child.nextSibling) stack.push(child);
	}
	return used;
}

/** Allocates unused `rIdN` ids for new relationships introduced while applying a model diff. */
export class RelationshipAllocator {
	private readonly used: Set<string>;
	private readonly created: NewRelationship[] = [];
	private readonly binaryPartIds = new Map<string, string>();

	/** Shared drawing-id allocator, so `wp:docPr` ids stay unique across every part of a save. */
	readonly docPrIds: DocPrIdAllocator;

	/** The ink parts this part's relationships name, by relationship id. */
	readonly inkParts = new Map<string, string>();

	constructor(existingIds: Iterable<string>, docPrIds: DocPrIdAllocator = new DocPrIdAllocator()) {
		this.used = new Set(existingIds);
		this.docPrIds = docPrIds;
	}

	private nextId(): string {
		const id = nextRelationshipId(this.used);
		this.used.add(id);
		return id;
	}

	/** Registers a brand-new external hyperlink target and returns its relationship id. */
	addExternalHyperlink(href: string): string {
		if (!isWordNavigableHref(href))
			throw new Error(
				`Unsupported hyperlink target: only http(s), mailto, tel and ftp links are supported.`,
			);
		const id = this.nextId();
		this.created.push({ id, type: HYPERLINK_RELATIONSHIP_TYPE, target: href, mode: 'External' });
		return id;
	}

	/** Registers (once per part) a newly inserted image part and returns its relationship id. */
	addImage(partName: string): string {
		return this.addBinaryPart(IMAGE_RELATIONSHIP_TYPE, partName);
	}

	addBinaryPart(type: string, partName: string): string {
		const key = `${type}:${partName}`;
		const existing = this.binaryPartIds.get(key);
		if (existing) return existing;
		const id = this.nextId();
		this.created.push({
			id,
			type,
			target: partName.replace(/^word\//, ''),
			mode: 'Internal',
			partName,
		});
		this.binaryPartIds.set(key, id);
		return id;
	}

	/** Registers a part of a newly inserted diagram and returns its relationship id. */
	addDiagramPart(part: NewDiagramPart): string {
		const id = this.nextId();
		this.created.push({
			id,
			type: part.type,
			target: part.partName.replace(/^word\//, ''),
			mode: 'Internal',
			partName: part.partName,
			overrideContentType: part.contentType,
			...(part.content !== undefined ? { content: part.content } : {}),
		});
		return id;
	}

	/** Registers a chart inserted in this session; its part is named when the relationships are written. */
	addChart(space: ChartSpace): string {
		const id = this.nextId();
		this.created.push({
			id,
			type: CHART_RELATIONSHIP_TYPE,
			target: '',
			mode: 'Internal',
			chart: space,
		});
		return id;
	}

	/** Registers an extended chart inserted in this session; its part is named on write. */
	addChartEx(space: ChartExSpace): string {
		const id = this.nextId();
		this.created.push({
			id,
			type: CHART_EX_RELATIONSHIP_TYPE,
			target: '',
			mode: 'Internal',
			chartEx: space,
		});
		return id;
	}

	/** Registers ink drawn in this session; its InkML part is named and written on save. */
	addInk(ink: DocxInk): string {
		const id = this.nextId();
		this.created.push({
			id,
			type: INK_RELATIONSHIP,
			target: '',
			mode: 'Internal',
			ink,
		});
		return id;
	}

	get newRelationships(): readonly NewRelationship[] {
		return this.created;
	}
}
