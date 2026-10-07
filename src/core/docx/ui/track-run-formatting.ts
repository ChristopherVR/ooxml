import { propertiesSignature } from '../revision-properties';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { AddMarkStep, RemoveMarkStep, AttrStep, type Step } from 'prosemirror-transform';
import type { TextRun } from '../model';
import { createRun } from '../write-run';
import { buildXml, first, makeW, parseXml, WORD_NS, type XmlElement } from '../xml';
import { parseRunPropertiesSnapshot } from '../restore-run-format';
import { inlineNodeRun } from './run-adapter';
import { formattingRevision } from './review-formatting';
import {
	RUN_FORMAT_MARKS as TRACKED_MARKS,
	setInlineRunFormatting,
	hasInlineRunAttributes,
} from './inline-formatting';

/** Uses the ordinary writer so snapshots retain script fonts and opaque source properties. */
function properties(node: ProseMirrorNode): XmlElement {
	const run = inlineNodeRun(node) ?? { text: '' };
	// Only serialize properties here. The drawing/OMML has its own source and needs no allocator.
	run.text = '';
	delete run.image;
	delete run.equation;
	delete run.revision;
	delete run.formatRevision;
	const doc = parseXml(`<w:r xmlns:w="${WORD_NS}"/>`);
	return first(createRun(doc, run), 'rPr') ?? makeW(doc, 'rPr');
}

/** Records pure, supported formatting transactions without changing text or inline identities. */
export function trackRunFormatting(
	steps: readonly Step[],
	oldState: EditorState,
	newState: EditorState,
	author: string,
	date: string,
	nextId: () => string,
): Transaction | null {
	const markType = newState.schema.marks.runProperties;
	if (
		!markType ||
		!steps.length ||
		!steps.every(
			(step) =>
				((step instanceof AddMarkStep || step instanceof RemoveMarkStep) &&
					TRACKED_MARKS.has(step.mark.type.name)) ||
				(step instanceof AttrStep && step.attr === 'format'),
		)
	)
		return null;
	const ranges = (steps as (AddMarkStep | RemoveMarkStep | AttrStep)[])
		.map((step) =>
			step instanceof AttrStep
				? { from: step.pos, to: step.pos + 1 }
				: { from: step.from, to: step.to },
		)
		.sort((a, b) => a.from - b.from);
	const merged: { from: number; to: number }[] = [];
	for (const range of ranges) {
		const last = merged.at(-1);
		if (last && range.from <= last.to) last.to = Math.max(last.to, range.to);
		else merged.push({ ...range });
	}
	const tr = newState.tr;
	let id: string | undefined;
	for (const range of merged)
		newState.doc.nodesBetween(range.from, range.to, (node, pos) => {
			if (!node.isInline || node.type.name === 'equation') return;
			const from = Math.max(range.from, pos);
			const to = Math.min(range.to, pos + node.nodeSize);
			// The old text can contain several independently formatted runs inside one new node.
			oldState.doc.nodesBetween(from, to, (previous, previousPos) => {
				if (!previous.isInline || previous.type.name === 'equation') return;
				const start = Math.max(from, previousPos);
				const end = Math.min(to, previousPos + previous.nodeSize);
				const before = properties(previous);
				const after = properties(node);
				if (propertiesSignature(before) === propertiesSignature(after)) {
					if (
						hasInlineRunAttributes(node) &&
						node.marks.some((mark) => TRACKED_MARKS.has(mark.type.name))
					)
						setInlineRunFormatting(tr, start, inlineNodeRun(node)!);
					return;
				}
				const pending = formattingRevision(previous);
				const priorXml = pending?.previousRunPropertiesXml ?? buildXml(before);
				const restored =
					propertiesSignature(after) === propertiesSignature(parseRunPropertiesSnapshot(priorXml));
				const mark = node.marks.find((item) => item.type === markType);
				const atom = hasInlineRunAttributes(node);
				const props = atom
					? inlineNodeRun(node)!
					: (structuredClone(mark?.attrs.props ?? {}) as Partial<TextRun>);
				delete props.formatRevision;
				if (props.revision?.kind === 'formatChange') delete props.revision;
				if (!restored)
					props.formatRevision = pending ?? {
						kind: 'formatChange',
						author,
						date,
						dateUtc: date,
						id: (id ??= nextId()),
						previousRunPropertiesXml: priorXml,
					};
				if (atom) {
					setInlineRunFormatting(tr, start, props as TextRun);
					return;
				}
				tr.removeMark(start, end, markType);
				if (Object.keys(props).length) tr.addMark(start, end, markType.create({ props }));
			});
		});
	return tr.docChanged ? tr : null;
}
