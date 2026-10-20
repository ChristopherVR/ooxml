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
import type { ContentControl } from './content-control-model';
import { inlineSourceLookup } from './inline-source';
import { takeSimpleFieldCommentEdges } from './simple-field-comments';
import { writeFieldFlags } from './field-flags';
import { sameFieldFlags } from './field-lock';
import { isEquationElement, writeEquation } from './equation';
import type { PlacedRun } from './inline-markers';
import type { InlineSlot } from './inline-slots';
import { convertToDeleteText, revisionWrapper } from './write-revisions';
import { getW, makeW, REL_NS, WORD_NS, type XmlDocument, type XmlElement } from './xml';

function sameLink(a: HyperlinkInfo | undefined, b: HyperlinkInfo | undefined): boolean {
	if (!a && !b) return true;
	if (!a || !b) return false;
	return a.href === b.href && a.anchor === b.anchor && a.tooltip === b.tooltip;
}

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
	if (wrapped) {
		if (revision.kind === 'delete') convertToDeleteText(doc, node);
		nodes.push(revisionWrapper(doc, revision, node));
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
		placed: { ...(nodes[0] && { first: nodes[0] }), ...(text && !wrapped && { splittable: node }) },
	};
}

/**
 * Builds the paragraph's ordered top-level inline nodes from the model's flat run list, and where
 * each run was written (inside its hyperlink or content control when it has one).
 */
export function buildInlineContent(
	doc: XmlDocument,
	runs: TextRun[],
	base: TextRun[] | undefined,
	slots: InlineSlot[],
	allocator?: RelationshipAllocator,
	ranges?: ParagraphRanges,
): { nodes: XmlElement[]; placed: PlacedRun[] } {
	// Ranges continuing from an earlier paragraph or into a later one open or close there instead.
	const { opens, closes } = rangeEdges(runs, ranges);
	const runAt = (at: number) => expectDefined(runs[at], `run ${at}`);
	const sourceFor = inlineSourceLookup(runs, base);
	const placed: PlacedRun[] = [];
	const nodesFor = (index: number) => {
		const written = runNodesAt(index);
		placed[index] = written.placed;
		return written.nodes;
	};
	const runNodesAt = (index: number) => {
		const run = runAt(index);
		const sourceIndex = sourceFor(index);
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
		const oldContainers = new Set(slots.slice(start, index).map((slot) => slot?.container));
		const [onlyContainer] = oldContainers;
		const reusable =
			oldContainers.size === 1 &&
			onlyContainer !== undefined &&
			onlyContainer.localName ===
				(group.kind === 'link' ? 'hyperlink' : group.kind === 'control' ? 'sdt' : 'fldSimple') &&
			base?.slice(start, index).every((run) => sameGroup(run && containerGroup(run), group)) ===
				true;
		let container: XmlElement;
		let target: XmlElement | undefined;
		if (reusable && group.kind === 'control') {
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
		for (let item = start; item < index; item++)
			for (const node of nodesFor(item)) (target ?? container).appendChild(node);
		output.push(...edges.before, container, ...edges.after);
	}
	return { nodes: output, placed };
}

type ContainerGroup =
	| { kind: 'link'; link: HyperlinkInfo }
	| { kind: 'control'; control: ContentControl }
	| {
			kind: 'simpleField';
			instr: string;
			instanceId: string | undefined;
			fieldFlags: TextRun['fieldFlags'];
	  };

/** The wrapper a run is written inside: a hyperlink, or a simple field (`w:fldSimple`). */
function containerGroup(run: TextRun): ContainerGroup | undefined {
	if (run.contentControl) return { kind: 'control', control: run.contentControl };
	if (run.link) return { kind: 'link', link: run.link };
	if (run.field?.simple)
		return {
			kind: 'simpleField',
			instr: run.field.instr,
			instanceId: run.fieldInstanceId,
			fieldFlags: run.fieldFlags,
		};
	return undefined;
}

function sameGroup(a: ContainerGroup | undefined, b: ContainerGroup | undefined): boolean {
	if (!a || !b) return a === b;
	if (a.kind === 'control' || b.kind === 'control')
		return a.kind === 'control' && b.kind === 'control' && a.control.id === b.control.id;
	if (a.kind === 'link' && b.kind === 'link') return sameLink(a.link, b.link);
	return (
		a.kind === 'simpleField' &&
		b.kind === 'simpleField' &&
		a.instr === b.instr &&
		a.instanceId === b.instanceId &&
		sameFieldFlags(a.fieldFlags, b.fieldFlags)
	);
}
