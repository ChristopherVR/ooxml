// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Rebuilds a paragraph's inline content (runs, pictures, hyperlinks, tracked-change wrappers and
// comment/move range markers) from the model's flat run list: hyperlink and field grouping plus
// the revision wrappers from write-revisions.ts and range markers from write-ranges.ts.
import { expectDefined } from './expect-defined';
import type { HyperlinkInfo, TextRun } from './model';
import type { RelationshipAllocator } from './relationship-allocator';
import {
	rangeEdges,
	rangeEndNodes,
	rangeStartNodes,
	type ParagraphRanges,
	type RangeEdge,
} from './write-ranges';

import { createRun } from './write-run';
import { createContentControl, resetContentControl } from './write-content-control';
import { inlineSourceLookup } from './inline-source';
import { takeSimpleFieldCommentEdges } from './simple-field-comments';
import { writeFieldFlags } from './field-flags';
import { isEquationElement, writeEquation } from './equation';
import { containerGroup, containerSources, sameGroup } from './inline-containers';
import type { PlacedRun } from './inline-markers';
import type { InlineSlot } from './inline-slots';
import { convertToDeleteText, revisionWrapper } from './write-revisions';
import { getW, makeW, REL_NS, WORD_NS, type XmlDocument, type XmlElement } from './xml';

function setW(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}

function newHyperlink(
	doc: XmlDocument,
	link: HyperlinkInfo,
	allocator?: RelationshipAllocator,
): XmlElement {
	const container = makeW(doc, 'hyperlink');
	if (link.href) {
		if (!allocator)
			throw new Error('A relationship allocator is required to write new hyperlinks.');
		container.setAttributeNS(REL_NS, 'r:id', allocator.addExternalHyperlink(link.href));
	}
	if (link.anchor) setW(container, 'anchor', link.anchor);
	if (link.tooltip) setW(container, 'tooltip', link.tooltip);
	if (!getW(container, 'history')) setW(container, 'history', '1');
	return container;
}

/** Builds one run's nodes: its range starts, the (revision-wrapped) run, and range ends. */
function runNodes(
	doc: XmlDocument,
	run: TextRun,
	base: TextRun | undefined,
	old: XmlElement | undefined,
	opens: RangeEdge[],
	closes: RangeEdge[],
	allocator?: RelationshipAllocator,
): { nodes: XmlElement[]; placed: PlacedRun } {
	// Comments open outside moves and close outside them, so ranges nest.
	const nodes = opens.flatMap((edge) => rangeStartNodes(doc, edge));
	const node = run.equation
		? writeEquation(doc, run, base, old)
		: createRun(doc, run, base, old, allocator);
	const revision = run.revision;
	const wrapped = revision && ['insert', 'delete', 'moveFrom', 'moveTo'].includes(revision.kind);
	let wrapper: XmlElement | undefined;
	if (wrapped) {
		if (revision.kind === 'delete') convertToDeleteText(doc, node);
		wrapper = revisionWrapper(doc, revision, node);
		nodes.push(wrapper);
	} else nodes.push(node);
	for (const edge of [...closes].reverse()) nodes.push(...rangeEndNodes(doc, edge));
	const text =
		!run.equation &&
		!run.image &&
		!run.fieldChar &&
		run.fieldCode === undefined &&
		!run.noteReference &&
		!run.noteMark &&
		!run.noteSeparator &&
		!run.break;
	return {
		nodes,
		placed: {
			...(nodes[0] && { first: nodes[0], last: nodes.at(-1)! }),
			...(text && !wrapped && { splittable: node }),
			...(wrapper && { wrapper }),
		},
	};
}

/**
 * Builds the paragraph's ordered top-level inline nodes from the model's flat run list, where
 * each run was written (inside its hyperlink or content control when it has one), and what was
 * written in place of each source hyperlink, simple field or content control (`containers`: the
 * elements its runs went into, the control's content for a control).
 */
export function buildInlineContent(
	doc: XmlDocument,
	runs: TextRun[],
	base: TextRun[] | undefined,
	slots: InlineSlot[],
	allocator?: RelationshipAllocator,
	ranges?: ParagraphRanges,
	sources = containerSources(runs, base, slots),
): { nodes: XmlElement[]; placed: PlacedRun[]; containers: Map<XmlElement, Set<XmlElement>> } {
	// Ranges continuing from an earlier paragraph or into a later one open or close there instead.
	const { opens, closes } = rangeEdges(runs, ranges);
	const runAt = (at: number) => expectDefined(runs[at], `run ${at}`);
	const sourceFor = inlineSourceLookup(runs, base);
	const placed: PlacedRun[] = [];
	const containers = new Map<XmlElement, Set<XmlElement>>();
	const nodesFor = (index: number) => {
		// The source run is looked up once: picture lookups take their match.
		const sourceIndex = sourceFor(index);
		const written = runNodesAt(index, sourceIndex);
		// A run still in a tracked change of the kind it was read in continues that change.
		const revision = sourceIndex === undefined ? undefined : slots[sourceIndex]?.revision;
		const wrapper = written.placed.wrapper;
		placed[index] =
			revision && wrapper?.localName === revision.localName
				? { ...written.placed, revision }
				: written.placed;
		return written.nodes;
	};
	const runNodesAt = (index: number, sourceIndex: number | undefined) => {
		const run = runAt(index);
		const source = sourceIndex === undefined ? undefined : base?.[sourceIndex];
		const old = sourceIndex === undefined ? undefined : slots[sourceIndex]?.element;
		return runNodes(
			doc,
			run,
			(!run.equation && source?.equation) || (!run.image && source?.image) ? undefined : source,
			(!run.equation && old && isEquationElement(old)) || (!run.image && source?.image)
				? undefined
				: old,
			opens.get(index) ?? [],
			closes.get(index) ?? [],
			allocator,
		);
	};
	const output: XmlElement[] = [];
	let index = 0;
	while (index < runs.length) {
		const group = containerGroup(runAt(index));
		if (!group) {
			output.push(...nodesFor(index));
			index++;
			continue;
		}
		const start = index;
		while (index < runs.length && sameGroup(containerGroup(runAt(index)), group)) index++;
		// The source container this group continues (see inline-containers.ts).
		const source = sources.get(start);
		const oldContainers = source?.containers ?? new Set<XmlElement | undefined>();
		const [onlyContainer] = oldContainers;
		const reusable =
			source?.same === true &&
			oldContainers.size === 1 &&
			onlyContainer !== undefined &&
			onlyContainer.localName ===
				(group.kind === 'link' ? 'hyperlink' : group.kind === 'control' ? 'sdt' : 'fldSimple');
		let container: XmlElement;
		let target: XmlElement | undefined;
		if (reusable && source.clone) container = onlyContainer.cloneNode(false) as XmlElement;
		else if (reusable && group.kind === 'control') {
			container = onlyContainer;
			target = resetContentControl(doc, container, group.control);
		} else if (reusable) {
			container = onlyContainer;
			for (const child of Array.from(container.childNodes)) container.removeChild(child);
		} else if (group.kind === 'control') {
			const created = createContentControl(doc, group.control);
			container = created.sdt;
			target = created.content;
		} else if (group.kind === 'link') container = newHyperlink(doc, group.link, allocator);
		else {
			container = makeW(doc, 'fldSimple');
			setW(container, 'instr', ` ${group.instr} `);
		}
		if (group.kind === 'simpleField') writeFieldFlags(container, runAt(start).fieldFlags);
		const edges =
			group.kind === 'simpleField'
				? takeSimpleFieldCommentEdges(doc, runs, start, index, opens, closes)
				: { before: [], after: [] };
		const parent = target ?? container;
		if (oldContainers.size === 1 && onlyContainer)
			containers.set(onlyContainer, (containers.get(onlyContainer) ?? new Set()).add(parent));
		for (let item = start; item < index; item++) {
			for (const node of nodesFor(item)) parent.appendChild(node);
			placed[item] = { ...placed[item], parent };
		}
		output.push(...edges.before, container, ...edges.after);
	}
	return { nodes: output, placed, containers };
}
