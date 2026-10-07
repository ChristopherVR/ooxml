import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { AddMarkStep, RemoveMarkStep, type Step } from 'prosemirror-transform';
import type { TextRun } from '../model';
import { createRun } from '../write-run';
import { buildXml, first, makeW, parseXml, WORD_NS, type XmlElement } from '../xml';
import { parseRunPropertiesSnapshot } from '../restore-run-format';
import { applyMarkFormatting } from './run-mark-properties';
import { formattingRevision } from './review-formatting';

const TRACKED_MARKS = new Set([
	'bold',
	'italic',
	'underline',
	'strike',
	'highlight',
	'verticalAlign',
	'language',
	'runRtl',
	'font',
	'characterStyle',
	'runProperties',
]);

/** Uses the ordinary writer so snapshots retain script fonts and opaque source properties. */
function properties(node: ProseMirrorNode): XmlElement {
	const run: TextRun = { text: '' };
	applyMarkFormatting(run, node);
	delete run.revision;
	delete run.formatRevision;
	const doc = parseXml(`<w:r xmlns:w="${WORD_NS}"/>`);
	return first(createRun(doc, run), 'rPr') ?? makeW(doc, 'rPr');
}

/** Namespace-aware comparison ignores serialization prefixes and attribute ordering. */
function signature(element: XmlElement): string {
	return JSON.stringify([
		element.namespaceURI,
		element.localName,
		Array.from(element.attributes)
			.filter((attribute) => attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
			.map((attribute) => [attribute.namespaceURI, attribute.localName, attribute.value])
			.sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
		Array.from(element.childNodes).flatMap((child) =>
			child.nodeType === 1
				? [signature(child as XmlElement)]
				: child.textContent?.trim()
					? [child.textContent]
					: [],
		),
	]);
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
				(step instanceof AddMarkStep || step instanceof RemoveMarkStep) &&
				TRACKED_MARKS.has(step.mark.type.name),
		)
	)
		return null;
	const ranges = (steps as (AddMarkStep | RemoveMarkStep)[])
		.map(({ from, to }) => ({ from, to }))
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
			if (!node.isText && node.type.name !== 'hardBreak') return;
			const from = Math.max(range.from, pos);
			const to = Math.min(range.to, pos + node.nodeSize);
			// The old text can contain several independently formatted runs inside one new node.
			oldState.doc.nodesBetween(from, to, (previous, previousPos) => {
				if (!previous.isText && previous.type.name !== 'hardBreak') return;
				const start = Math.max(from, previousPos);
				const end = Math.min(to, previousPos + previous.nodeSize);
				const before = properties(previous);
				const after = properties(node);
				if (signature(before) === signature(after)) return;
				const pending = formattingRevision(previous);
				const priorXml = pending?.previousRunPropertiesXml ?? buildXml(before);
				const restored = signature(after) === signature(parseRunPropertiesSnapshot(priorXml));
				const mark = node.marks.find((item) => item.type === markType);
				const props = structuredClone(mark?.attrs.props ?? {}) as Partial<TextRun>;
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
				tr.removeMark(start, end, markType);
				if (Object.keys(props).length) tr.addMark(start, end, markType.create({ props }));
			});
		});
	return tr.docChanged ? tr : null;
}
