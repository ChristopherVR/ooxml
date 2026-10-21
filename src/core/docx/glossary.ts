// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The glossary document (`word/glossary/document.xml`) holds a document's building blocks, among
// them the placeholder text of its content controls (`w:sdtPr/w:placeholder/w:docPart`). It is
// read only to show those placeholders; the part itself is saved as it was.
import type JSZip from 'jszip';
import type { ContentControl } from './content-control-model';
import type { Block, DocumentModel } from './model';
import { resolveInternalTarget, type Relationship } from './package-parts';
import { children, first, getW, isElement, named, parseXml, type XmlElement } from './xml';

const GLOSSARY_DOCUMENT =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument';

/** The text a paragraph shows: its `w:t` text, with `w:tab` as a tab. */
function paragraphText(paragraph: XmlElement): string {
	let text = '';
	const visit = (node: XmlElement) => {
		for (const child of Array.from(node.childNodes).filter(isElement)) {
			if (named(child, 't')) text += child.textContent ?? '';
			else if (named(child, 'tab')) text += '\t';
			else if (!named(child, 'rPr') && !named(child, 'pPr')) visit(child);
		}
	};
	visit(paragraph);
	return text;
}

/** Each glossary entry's text by entry name: the first paragraph of its body. */
export function parseGlossaryTexts(xml: string): Map<string, string> {
	const texts = new Map<string, string>();
	const parts = first(parseXml(xml).documentElement, 'docParts');
	for (const part of parts ? children(parts, 'docPart') : []) {
		const name = getW(first(first(part, 'docPartPr'), 'name'), 'val');
		const paragraph = first(first(part, 'docPartBody'), 'p');
		if (name && paragraph) texts.set(name, paragraphText(paragraph));
	}
	return texts;
}

/** Every content control of `blocks`: around runs, paragraphs, tables, rows and cells. */
export function forEachContentControl(
	blocks: readonly Block[],
	visit: (control: ContentControl) => void,
): void {
	for (const block of blocks) {
		if (block.contentControl) visit(block.contentControl);
		if (block.type === 'paragraph') {
			for (const run of block.runs) if (run.contentControl) visit(run.contentControl);
			continue;
		}
		for (const row of block.rowProperties ?? []) if (row.contentControl) visit(row.contentControl);
		for (const row of block.rows)
			for (const cell of row) {
				if (cell.contentControl) visit(cell.contentControl);
				forEachContentControl(cell.paragraphs, visit);
				if (cell.sourceBlocks) forEachContentControl(cell.sourceBlocks, visit);
			}
	}
}

/**
 * Gives each control whose placeholder is a glossary entry that entry's text. The glossary is read
 * only when a control names an entry, and a part that cannot be read leaves the built-in prompts
 * (with a warning) rather than stopping the document from opening.
 */
export async function readPlaceholderTexts(
	zip: JSZip,
	relationships: ReadonlyMap<string, Relationship>,
	model: DocumentModel,
): Promise<void> {
	const stories = [
		model.blocks,
		...(model.sections ?? []).flatMap((section) =>
			[section.headers, section.footers].flatMap((slots) =>
				Object.values(slots ?? {}).map((content) => content?.blocks ?? []),
			),
		),
		...[...(model.footnotes ?? []), ...(model.endnotes ?? [])].map((note) => note.blocks),
	];
	const named: ContentControl[] = [];
	for (const blocks of stories)
		forEachContentControl(blocks, (control) => {
			if (control.placeholder !== undefined) named.push(control);
		});
	if (!named.length) return;
	const relationship = [...relationships.values()].find(
		(entry) => entry.type === GLOSSARY_DOCUMENT && entry.mode === 'Internal',
	);
	const part = relationship && zip.file(resolveInternalTarget(relationship.target));
	if (!part) return;
	let texts: Map<string, string>;
	try {
		texts = parseGlossaryTexts(await part.async('string'));
	} catch (error) {
		model.warnings.push(
			`The glossary part could not be read (${error instanceof Error ? error.message : String(error)}), so content controls show the built-in placeholder prompts; the part is saved as it was.`,
		);
		return;
	}
	for (const control of named) {
		const text = texts.get(control.placeholder!);
		if (text !== undefined) control.placeholderText = text;
	}
}
