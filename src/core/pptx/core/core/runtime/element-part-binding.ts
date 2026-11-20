/**
 * Save-time part ownership for slide elements.
 *
 * Every copy path (duplicate, paste, a peer's edit, an undo snapshot, an SDK clone) copies an
 * element's model, and with it the ids and paths of the package parts it shows. Two things then
 * go wrong in the package unless the save repairs them:
 *
 * - The ids belong to the part the element was read from (`relationshipSourcePath`). On another
 *   slide they name that slide's own relationships, or none.
 * - A part one element owns (a chart, a diagram, an OLE object, see `ownsRelationshipTarget`) is
 *   now claimed by two, so an edit to one silently rewrites the other, and PowerPoint refuses the
 *   file.
 *
 * {@link bindElementParts} walks the deck in order. The first element to claim an owned part keeps
 * it; every later claimant gets its own copy (fresh names, the owned parts below it copied too).
 * An element on a part other than its source gets relationships of its own on that part. Shared
 * targets (images, media, hyperlinks) stay shared. Each element then records the slide as its
 * source, so the next save finds it at home.
 *
 * @module core/core/runtime/element-part-binding
 */
import type JSZip from 'jszip';
import { withDiagramDrawingRelationship } from '../../../../diagram/write/cache';
import { clonePart, relativePartTarget } from '../../../../opc/clone-part-graph';
import { ownsRelationshipTarget } from '../../../../opc/part-ownership';
import {
	buildRelationshipsXml,
	nextRelationshipId,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
	type Relationship,
} from '../../../../opc/relationships';
import type { PptxElement, PptxSlide } from '../../types';
import { cloneElement } from '../../utils/clone-utils';
import { isTemplateElementId } from '../../utils/group-ops';
import { remapPartPaths } from '../builders/duplicate-slide-parts';
import { elementRelationshipIds, renameElementRelationshipIds } from './element-relationship-ids';

/** Relationships of one part, keyed by id. */
export type PartRelationships = Map<string, Relationship>;

export interface ElementPartBindingInput {
	zip: JSZip;
	slides: readonly PptxSlide[];
	/**
	 * Relationships of the elements' source parts as they were before this save changed the
	 * package (a slide deleted this save has lost its own). Other parts are read from the package.
	 */
	sourceRelationships: ReadonlyMap<string, PartRelationships>;
	/** Which relationships of a copied part lead to parts the copy needs its own of. */
	shouldClone: (relationship: Relationship) => boolean;
	/**
	 * Moves session state keyed by part path (chart baselines, revisions) onto the copies, whose
	 * element now names them by the renamed relationship ids.
	 */
	onPartsCopied: (
		copies: ReadonlyMap<string, string>,
		renames: ReadonlyMap<string, string>,
	) => void;
	/** A part that could not be copied stays shared, and is reported. */
	onCopyFailed: (element: PptxElement, part: string, error: unknown) => void;
}

/**
 * Every element on a slide, group members included, in paint order. Layout and master elements
 * shown on the slide (`layout-` / `master-` ids) are saved into their own part and are skipped.
 */
function* allElements(elements: readonly PptxElement[]): Generator<PptxElement> {
	for (const element of elements) {
		if (isTemplateElementId(element.id)) continue;
		yield element;
		if (element.type === 'group') yield* allElements(element.children);
	}
}

/** See {@link ElementPartBindingInput.sourceRelationships}; read before slides are reconciled. */
export async function readElementSourceRelationships(
	zip: JSZip,
	slides: readonly PptxSlide[],
): Promise<Map<string, PartRelationships>> {
	const result = new Map<string, PartRelationships>();
	for (const slide of slides)
		for (const element of allElements(slide.elements)) {
			const source = element.relationshipSourcePath;
			if (source && !result.has(source))
				result.set(
					source,
					parseRelationships(await zip.file(relationshipsPartFor(source))?.async('string')),
				);
		}
	return result;
}

/**
 * The id of a relationship of `part` to `target` (a part path, or an address when external),
 * adding one when there is none.
 */
function relationshipTo(
	relationships: PartRelationships,
	part: string,
	{ type, mode }: Relationship,
	target: string,
): string {
	const internal = mode !== 'External';
	for (const relationship of relationships.values()) {
		const at = internal ? resolvePartPath(part, relationship.target) : relationship.target;
		if (relationship.type === type && relationship.mode === mode && at === target)
			return relationship.id;
	}
	const id = nextRelationshipId(new Set(relationships.keys()));
	relationships.set(id, {
		id,
		type,
		mode,
		target: internal ? relativePartTarget(part, target) : target,
	});
	return id;
}

interface Binding {
	renames: Map<string, string>;
	copies: Map<string, string>;
}

/** Point the element's own fields at its copies and at the slide's relationship ids. */
async function retargetElement(
	zip: JSZip,
	element: PptxElement,
	slidePath: string,
	source: string,
	{ renames, copies }: Binding,
	relationships: PartRelationships,
): Promise<void> {
	const drawingId = element.type === 'smartArt' ? element.smartArtData?.drawingRelId : undefined;
	Object.assign(element, cloneElement(element));
	renameElementRelationshipIds(element, renames);
	remapPartPaths(element, copies);
	if (element.type === 'ole' && element.oleTarget && !element.isLinked) {
		const part = resolvePartPath(source, element.oleTarget);
		element.oleTarget = relativePartTarget(slidePath, copies.get(part) ?? part);
	}
	// A diagram's data part names its drawing by the host slide's relationship id.
	const data = element.type === 'smartArt' ? element.smartArtData : undefined;
	if (!drawingId || !renames.has(drawingId) || !data?.dataRelId || !data.drawingRelId) return;
	const dataRelationship = relationships.get(data.dataRelId);
	if (!dataRelationship) return;
	const dataPart = resolvePartPath(slidePath, dataRelationship.target);
	const xml = await zip.file(dataPart)?.async('string');
	if (xml) zip.file(dataPart, withDiagramDrawingRelationship(xml, data.drawingRelId));
}

/** See the module comment. Changes slide relationships and elements in place. */
export async function bindElementParts(input: ElementPartBindingInput): Promise<void> {
	const { zip } = input;
	const owners = new Map<string, PptxElement>();
	const readRelationships = async (part: string): Promise<PartRelationships> =>
		input.sourceRelationships.get(part) ??
		parseRelationships(await zip.file(relationshipsPartFor(part))?.async('string'));

	/** Give `element` its own copy of `part`, returning the copy's path (or `part` on failure). */
	const copyOwnedPart = async (
		element: PptxElement,
		part: string,
		copies: Map<string, string>,
	): Promise<string> => {
		try {
			const copied = await clonePart({
				zip,
				part,
				copies: new Map(copies),
				shouldClone: input.shouldClone,
			});
			for (const [from, to] of copied) copies.set(from, to);
			return copied.get(part) ?? part;
		} catch (error) {
			input.onCopyFailed(element, part, error);
			return part;
		}
	};

	for (const slide of input.slides) {
		const relsPath = relationshipsPartFor(slide.id);
		const relationships = parseRelationships(await zip.file(relsPath)?.async('string'));
		let changed = false;
		for (const element of allElements(slide.elements)) {
			const source = element.relationshipSourcePath ?? slide.id;
			const declared = source === slide.id ? relationships : await readRelationships(source);
			const ids = elementRelationshipIds(element);
			const binding: Binding = { renames: new Map(), copies: new Map() };
			for (const id of ids) {
				const relationship = declared.get(id);
				if (!relationship) continue;
				const external = relationship.mode === 'External';
				let target = external ? relationship.target : resolvePartPath(source, relationship.target);
				let moved = source !== slide.id;
				if (!external && ownsRelationshipTarget(slide.id, relationship)) {
					const owner = owners.get(target);
					if (owner && owner !== element) {
						const copy = await copyOwnedPart(element, target, binding.copies);
						moved ||= copy !== target;
						target = copy;
					}
					owners.set(target, element);
				}
				if (!moved) continue;
				const newId = relationshipTo(relationships, slide.id, relationship, target);
				if (newId !== id) binding.renames.set(id, newId);
				changed = true;
			}
			if (binding.renames.size > 0 || binding.copies.size > 0) {
				await retargetElement(zip, element, slide.id, source, binding, relationships);
				if (binding.copies.size > 0) input.onPartsCopied(binding.copies, binding.renames);
			}
			if (ids.size > 0 && element.relationshipSourcePath !== slide.id)
				element.relationshipSourcePath = slide.id;
		}
		if (changed) zip.file(relsPath, buildRelationshipsXml(relationships));
	}
}
