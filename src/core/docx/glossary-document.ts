// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The document's own building blocks: the entries of its glossary document
// (`word/glossary/document.xml`, `w:docPart`). Quick Parts, AutoText and the Building Blocks
// Organizer read them, insert them and add new ones from a selection. Entries from the file are
// kept as they were (only entries added or deleted here change the part); pictures inside an entry
// are not resolved, since the glossary part has its own relationships.
import type JSZip from 'jszip';
import { createOfficeGuid } from '../crypto/uuid';
import type { Block, DocumentModel, Paragraph } from './model';
import { parseBlocksFromContainer } from './block-parser';
import { declaredEntrySize } from '../opc/index';
import { resolveInternalTarget, type Relationship } from './package-parts';
import { RelationshipAllocator } from './relationship-allocator';
import { applyBlocks } from './write';
import { buildXml, children, first, getW, makeW, parseXml, WORD_NS, type XmlElement } from './xml';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts';

const GLOSSARY_RELATIONSHIP =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument';
const GLOSSARY_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml';
export const GLOSSARY_PART = 'word/glossary/document.xml';
/** The glossary's own size limit: a part larger than this is left alone and not listed. */
const GLOSSARY_XML_LIMIT = 8_000_000;

/** Word's building block galleries that Quick Parts and AutoText show. */
export type GlossaryGallery = 'docParts' | 'autoTxt';

export interface GlossaryEntry {
	name: string;
	/** `w:docPartGallery`: `docParts` is the Quick Parts gallery, `autoTxt` AutoText; others are kept. */
	gallery: string;
	category: string;
	description?: string;
	blocks: Block[];
	/** Read from the file (and so already in the part); entries added in the editor lack it. */
	loaded?: true;
}

/** An entry's identity inside a glossary: Word keys building blocks by name, gallery and category. */
export const glossaryEntryKey = (entry: Pick<GlossaryEntry, 'name' | 'gallery' | 'category'>) =>
	`${entry.gallery}\u0000${entry.category}\u0000${entry.name}`;

function entryOf(part: XmlElement, index: number): GlossaryEntry | undefined {
	const properties = first(part, 'docPartPr');
	const name = getW(first(properties, 'name'), 'val');
	const body = first(part, 'docPartBody');
	if (!name || !body) return undefined;
	const category = first(properties, 'category');
	const description = getW(first(properties, 'description'), 'val');
	return {
		name,
		gallery: getW(first(category, 'gallery'), 'val') ?? 'placeholder',
		category: getW(first(category, 'name'), 'val') ?? 'General',
		...(description ? { description } : {}),
		blocks: parseBlocksFromContainer(body, `gl${index}-`),
		loaded: true,
	};
}

/** The glossary document's entries; a part that cannot be read yields none. */
export async function readGlossaryEntries(
	zip: JSZip,
	relationships: ReadonlyMap<string, Relationship>,
	model: DocumentModel,
): Promise<void> {
	const relationship = [...relationships.values()].find(
		(entry) => entry.type === GLOSSARY_RELATIONSHIP && entry.mode === 'Internal',
	);
	const partName = relationship && resolveInternalTarget(relationship.target);
	const file = partName && zip.file(partName);
	if (!file) return;
	model.glossaryPart = partName;
	// The size is checked from the zip directory, before the part is inflated.
	if (declaredEntrySize(file) > GLOSSARY_XML_LIMIT) return;
	try {
		const xml = await file.async('string');
		if (xml.length > GLOSSARY_XML_LIMIT) throw new Error('the part is too large');
		const parts = first(parseXml(xml).documentElement, 'docParts');
		const entries = (parts ? children(parts, 'docPart') : []).flatMap((part, index) => {
			const entry = entryOf(part, index);
			return entry ? [entry] : [];
		});
		if (entries.length) model.glossary = entries;
	} catch {
		// An unreadable part is left as it is and its entries are not listed; `readPlaceholderTexts`
		// is the one that reports a glossary it needs.
	}
}

const textOf = (paragraph: Paragraph): string => paragraph.runs.map((run) => run.text).join('');

/** The first line of an entry, for the Organizer's preview. */
export const glossaryPreview = (entry: GlossaryEntry): string =>
	entry.blocks.flatMap((block) => (block.type === 'paragraph' ? [textOf(block)] : []))[0] ?? '';

/**
 * Whether `blocks` can be written into a glossary entry: no pictures or drawings, links, notes or
 * comments, since those need relationships or parts the glossary part does not have.
 */
export function glossaryBlocksWritable(blocks: readonly Block[]): boolean {
	const paragraphs = (list: readonly Block[]): Paragraph[] =>
		list.flatMap((block) =>
			block.type === 'paragraph' ? [block] : block.rows.flat().flatMap((cell) => cell.paragraphs),
		);
	return paragraphs(blocks).every((paragraph) =>
		paragraph.runs.every(
			(run) => !run.image && !run.link && !run.noteReference && !run.commentIds?.length,
		),
	);
}

/** `blocks` with every paragraph given a new id, so an inserted copy never repeats one in use. */
export function glossaryBlocksCopy(blocks: readonly Block[], prefix: string): Block[] {
	let next = 0;
	const copy = structuredClone(blocks) as Block[];
	for (const block of copy) {
		if (block.type === 'paragraph') block.id = `${prefix}${next++}`;
		else {
			block.id = `${prefix}${next++}`;
			for (const cell of block.rows.flat())
				for (const p of cell.paragraphs) p.id = `${prefix}${next++}`;
		}
	}
	return copy;
}

function docPartElement(doc: XmlDocument, entry: GlossaryEntry): XmlElement {
	const part = makeW(doc, 'docPart');
	const properties = makeW(doc, 'docPartPr');
	const valued = (local: string, value: string) => {
		const element = makeW(doc, local);
		element.setAttributeNS(WORD_NS, 'w:val', value);
		return element;
	};
	properties.appendChild(valued('name', entry.name));
	const category = makeW(doc, 'category');
	category.appendChild(valued('name', entry.category));
	category.appendChild(valued('gallery', entry.gallery));
	properties.appendChild(category);
	const types = makeW(doc, 'types');
	types.appendChild(valued('type', entry.gallery === 'autoTxt' ? 'normal' : 'bbPlcHdr'));
	properties.appendChild(types);
	const behaviors = makeW(doc, 'behaviors');
	behaviors.appendChild(valued('behavior', 'content'));
	properties.appendChild(behaviors);
	if (entry.description) properties.appendChild(valued('description', entry.description));
	properties.appendChild(
		valued('guid', `{${createOfficeGuid().replace(/[{}]/g, '').toUpperCase()}}`),
	);
	part.appendChild(properties);
	const body = makeW(doc, 'docPartBody');
	const allocator = new RelationshipAllocator([]);
	applyBlocks(doc, body, entry.blocks, [], allocator, 9360);
	if (allocator.newRelationships.length)
		throw new Error(
			'A building block cannot hold links, notes or pictures: the glossary part has no relationships for them.',
		);
	part.appendChild(body);
	return part;
}

type XmlDocument = ReturnType<typeof parseXml>;

const EMPTY_GLOSSARY = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument xmlns:w="${WORD_NS}"><w:docParts/></w:glossaryDocument>`;

/**
 * Writes the entries added or deleted since the document was loaded: new `w:docPart`s are appended
 * (the part, its content type and its relationship are created when the document has none), and
 * deleted ones are removed. Every other entry and part is kept byte for byte.
 */
export async function writeGlossaryEntries(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel | undefined,
): Promise<void> {
	const now = model.glossary ?? [];
	const before = base?.glossary ?? [];
	const kept = new Set(now.filter((entry) => entry.loaded).map(glossaryEntryKey));
	const removed = before.filter((entry) => !kept.has(glossaryEntryKey(entry)));
	const added = now.filter((entry) => !entry.loaded);
	if (!removed.length && !added.length) return;
	const partName = model.glossaryPart ?? GLOSSARY_PART;
	const file = zip.file(partName);
	const doc = parseXml(file ? await file.async('string') : EMPTY_GLOSSARY);
	let parts = first(doc.documentElement, 'docParts');
	if (!parts) {
		parts = makeW(doc, 'docParts');
		doc.documentElement.appendChild(parts);
	}
	const dropped = new Set(removed.map(glossaryEntryKey));
	for (const part of children(parts, 'docPart')) {
		const entry = entryOf(part, 0);
		if (entry && dropped.has(glossaryEntryKey(entry))) parts.removeChild(part);
	}
	const present = new Set(
		children(parts, 'docPart').flatMap((part) => {
			const entry = entryOf(part, 0);
			return entry ? [glossaryEntryKey(entry)] : [];
		}),
	);
	for (const entry of added) {
		// Saving a selection under a name already in the gallery replaces that entry.
		if (present.has(glossaryEntryKey(entry)))
			for (const part of children(parts, 'docPart')) {
				const old = entryOf(part, 0);
				if (old && glossaryEntryKey(old) === glossaryEntryKey(entry)) parts.removeChild(part);
			}
		parts.appendChild(docPartElement(doc, entry));
	}
	zip.file(partName, buildXml(doc));
	if (!file && !model.glossaryPart) {
		await ensureContentTypeOverride(zip, GLOSSARY_PART, GLOSSARY_CONTENT_TYPE);
		await ensureDocumentRelationship(zip, GLOSSARY_RELATIONSHIP, 'glossary/document.xml');
	}
}
