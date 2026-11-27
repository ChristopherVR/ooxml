// Saving Word ink. Ink without its own drawing in the source (drawn in this session, pasted, moved
// to another paragraph or copied) gets a new InkML part (`word/ink/inkN.xml`), so no two drawings
// share a part. Kept ink whose strokes changed (erased, collected, Ink to Shape) is written back
// over its own part, keeping the nodes the model does not hold; ink only resized keeps its part,
// as Word stretches the traces onto the drawing's box. The parts of deleted ink are removed with
// their relationships. A part read only in part is never rewritten: a new drawing of it copies its
// bytes. The InkML writer loads on first use.
import type JSZip from 'jszip';
import { documentDrawingBlocks } from './chart';
import { INK_RELATIONSHIP, inksIn, type DocxInk } from './ink';
import { forEachParagraph } from './parse-warnings';
import type { DocumentModel, InlineImage } from './model';
import { buildRelationshipsXml, parseRelationships } from './package-parts';
import { relationshipsPartFor } from './part-relationships';
import { ensureContentTypeOverride, removeContentTypeOverride } from '../opc/index';

const INKML_CONTENT_TYPE = 'application/inkml+xml';

/** Writes the part of ink that has no part of its own in this save and returns its name. */
export async function writeNewInkPart(zip: JSZip, ink: DocxInk): Promise<string> {
	let index = 1;
	while (zip.file(`word/ink/ink${index}.xml`)) index++;
	const partName = `word/ink/ink${index}.xml`;
	const source = ink.partName ? zip.file(ink.partName) : null;
	const { updatedInkPart } = await import('#inkml-runtime');
	const xml = source && !ink.truncated ? await source.async('string') : undefined;
	// Unchanged (moved or copied) or read only in part: the source part's bytes, Word's encoding
	// included; otherwise the strokes as they are now.
	if (source && (ink.truncated || updatedInkPart(xml, ink.strokes, ink) === undefined))
		zip.file(partName, await source.async('uint8array'));
	else zip.file(partName, updatedInkPart(undefined, ink.strokes, ink)!);
	await ensureContentTypeOverride(zip, partName, INKML_CONTENT_TYPE);
	return partName;
}

function inkImages(blocks: DocumentModel['blocks']): InlineImage[] {
	const images: InlineImage[] = [];
	forEachParagraph(blocks, (paragraph) => {
		for (const run of paragraph.runs) if (run.image?.ink?.partName) images.push(run.image);
	});
	return images;
}

const partOf = (target: string) => `word/${target.replace(/^\/?word\//, '')}`;

/** The ink parts the saved body's drawings refer to (`w14:contentPart r:id`). */
async function bodyInkParts(zip: JSZip) {
	const relsPart = relationshipsPartFor('word/document.xml');
	const rels = parseRelationships(await zip.file(relsPart)?.async('string'));
	const xml = (await zip.file('word/document.xml')?.async('string')) ?? '';
	const parts = new Set<string>();
	for (const match of xml.matchAll(/<w14:contentPart\b[^>]*?\br:id="([^"]+)"/gu)) {
		const rel = rels.get(match[1] ?? '');
		if (rel?.type === INK_RELATIONSHIP) parts.add(partOf(rel.target));
	}
	return { parts, rels, relsPart };
}

/**
 * Rewrites the part of every ink that kept its drawing (`keptInk`) and whose strokes changed since `base`, and removes the parts of the
 * body's ink that the saved body no longer refers to (deleted, or moved to a part of its own),
 * with their relationships. Runs after the body is written.
 */
export async function applyInkPartEdits(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel | undefined,
	keptInk: ReadonlySet<object>,
): Promise<void> {
	if (!base) return;
	const bodyParts = new Set(inksIn(base.blocks).flatMap((ink) => ink.partName ?? []));
	const body = bodyParts.size ? await bodyInkParts(zip) : undefined;
	const loaded = new Map<string, string>();
	for (const ink of inksIn(documentDrawingBlocks(base)))
		if (ink.partName) loaded.set(ink.partName, JSON.stringify(ink));
	const changed = inkImages(documentDrawingBlocks(model)).filter((image) => {
		const ink = image.ink!;
		const before = loaded.get(ink.partName!);
		// Only ink that kept its own drawing: a copy still names the part it was copied from.
		return (
			keptInk.has(ink) && !ink.truncated && before !== undefined && before !== JSON.stringify(ink)
		);
	});
	if (changed.length) {
		const { updatedInkPart } = await import('#inkml-runtime');
		const saved = new Set<string>();
		for (const image of changed) {
			const ink = image.ink!;
			if (saved.has(ink.partName!)) continue;
			saved.add(ink.partName!);
			const source = await zip.file(ink.partName!)?.async('string');
			const xml = updatedInkPart(source, ink.strokes, ink);
			if (xml !== undefined) zip.file(ink.partName!, xml);
		}
	}
	if (!body) return;
	// A part a header, footer or note still shows stays.
	const elsewhere = new Set(
		inksIn(documentDrawingBlocks({ ...base, blocks: [] })).flatMap((ink) => ink.partName ?? []),
	);
	const removed = [...bodyParts].filter((part) => !body.parts.has(part) && !elsewhere.has(part));
	if (!removed.length) return;
	for (const [id, rel] of [...body.rels])
		if (rel.type === INK_RELATIONSHIP && removed.includes(partOf(rel.target))) body.rels.delete(id);
	zip.file(body.relsPart, buildRelationshipsXml(body.rels));
	for (const part of removed) {
		zip.remove(part);
		await removeContentTypeOverride(zip, part);
	}
}
