import type { Node } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { ReplaceStep, type Step } from 'prosemirror-transform';
import { parsePropertiesSnapshot, propertiesSignature } from '../revision-properties';
import { children, first } from '../xml';

function sourceProperties(value: unknown): unknown {
	if (typeof value !== 'string' || !value) return value;
	try {
		const properties = parsePropertiesSnapshot(value, 'pPr');
		const run = first(properties, 'rPr');
		if (run) {
			for (const name of ['ins', 'del'])
				for (const revision of children(run, name)) run.removeChild(revision);
			if (!run.childNodes.length && !run.attributes.length) properties.removeChild(run);
		}
		return propertiesSignature(properties);
	} catch {
		return value;
	}
}

function safeParagraph(node: Node, allowInsertion = false): boolean {
	if (node.type.name !== 'paragraph' || !node.type.spec.attrs?.markRevision) return false;
	if (node.attrs.markRevision && !allowInsertion) return false;
	let fields = false;
	node.descendants((child) => {
		if (child.type.name === 'fieldMarker' || child.marks.some((mark) => mark.type.name === 'field'))
			fields = true;
	});
	return !fields;
}

function endsSection(state: EditorState, node: Node): boolean {
	if (typeof state.doc.attrs.sections !== 'string') return false;
	try {
		const sections = JSON.parse(state.doc.attrs.sections) as { endsAtBlockId?: string }[];
		return sections.some(
			(section, index) =>
				section.endsAtBlockId === node.attrs.id &&
				(index < sections.length - 1 || state.doc.lastChild !== node),
		);
	} catch {
		return true;
	}
}

function sameProperties(first: Node, second: Node): boolean {
	return Object.keys({ ...first.attrs, ...second.attrs }).every(
		(key) =>
			key === 'id' ||
			key === 'markRevision' ||
			(key === 'sourceParagraphPropertiesXml'
				? sourceProperties(first.attrs[key]) === sourceProperties(second.attrs[key])
				: JSON.stringify(first.attrs[key]) === JSON.stringify(second.attrs[key])),
	);
}

/** Records one plain top-level paragraph split/join; complex structural edits remain unsupported. */
export function trackParagraphBoundary(
	steps: readonly Step[],
	oldState: EditorState,
	newState: EditorState,
	author: string,
	date: string,
	nextId: (kind: string) => string,
): Transaction | null {
	if (
		steps.length !== 1 ||
		!(steps[0] instanceof ReplaceStep) ||
		steps[0].toJSON().structure !== true
	)
		return null;
	const step = steps[0];
	const $from = oldState.doc.resolve(step.from);
	const $to = oldState.doc.resolve(step.to);
	if ($from.depth !== 1 || $to.depth !== 1) return null;
	const first = $from.parent;
	const pos = $from.before();
	if (endsSection(oldState, first)) return null;
	const revision = () => ({ author, date, dateUtc: date, id: nextId('revision') });
	if (step.from === step.to) {
		if (
			!safeParagraph(first) ||
			step.slice.openStart !== 1 ||
			step.slice.openEnd !== 1 ||
			step.slice.content.childCount !== 2 ||
			step.slice.content.child(0).content.size ||
			step.slice.content.child(1).content.size
		)
			return null;
		const left = newState.doc.nodeAt(pos);
		if (!left) return null;
		const rightPos = pos + left.nodeSize;
		const right = newState.doc.nodeAt(rightPos);
		if (
			!right ||
			!safeParagraph(left) ||
			!safeParagraph(right) ||
			!sameProperties(first, left) ||
			!sameProperties(first, right) ||
			!left.content.append(right.content).eq(first.content)
		)
			return null;
		const used = new Set<string>();
		newState.doc.descendants((node) => {
			if (node.attrs.id) used.add(String(node.attrs.id));
		});
		let id: string;
		do id = nextId('paragraph');
		while (used.has(id));
		// The final paragraph retains the original mark's formatting and identity,
		// so rejecting the inserted boundary restores the original paragraph.
		return newState.tr
			.setNodeMarkup(pos, undefined, {
				...left.attrs,
				id,
				markRevision: { ...revision(), kind: 'insert' },
			})
			.setNodeMarkup(rightPos, undefined, { ...first.attrs });
	}
	if (
		step.to !== step.from + 2 ||
		step.slice.size ||
		step.from !== $from.end() ||
		step.to !== $to.start()
	)
		return null;
	const second = $to.parent;
	// Different paragraph formatting requires Word's separate pPrChange history,
	// which is outside this plain-boundary recorder.
	if (!safeParagraph(first, true) || !safeParagraph(second) || !sameProperties(first, second))
		return null;
	const pending = first.attrs.markRevision;
	if (pending) {
		if (pending.kind !== 'insert' || pending.author !== author) return null;
		const merged = newState.doc.nodeAt(pos);
		if (!merged || !merged.content.eq(first.content.append(second.content))) return null;
		return newState.tr.setNodeMarkup(pos, undefined, { ...second.attrs });
	}
	// Keep both paragraph property snapshots while the deleted mark is pending.
	return newState.tr.step(step.invert(oldState.doc)).setNodeMarkup(pos, undefined, {
		...first.attrs,
		markRevision: { ...revision(), kind: 'delete' },
	});
}
