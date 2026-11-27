import type { Node as ProseMirrorNode, ResolvedPos } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import {
	AddMarkStep,
	AddNodeMarkStep,
	AttrStep,
	RemoveMarkStep,
	RemoveNodeMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
} from 'prosemirror-transform';
import { nodeControlOf } from './content-control-paste';
import { controlOfMark } from './content-control-spans';

const isTagged = (value: unknown): boolean =>
	!!value && typeof value === 'object' && 'appearance' in value && value.appearance === 'tags';

const controlMark = (node: ProseMirrorNode | null) =>
	node?.marks.find((mark) => mark.type.name === 'contentControl');

function holdsControl(node: ProseMirrorNode, tagsOnly: boolean): boolean {
	const mark = controlMark(node);
	const block = nodeControlOf(node)?.value;
	return tagsOnly ? isTagged(block) || isTagged(controlOfMark(mark)) : !!(block || mark);
}

/** A plain paragraph can separate two parts of one block control, even across a table. */
function besideBlockControl($pos: ResolvedPos, dir: -1 | 1): boolean {
	for (let depth = $pos.depth; depth >= 0; depth--) {
		const parent = $pos.node(depth);
		const start = dir < 0 ? $pos.index(depth) - 1 : $pos.indexAfter(depth);
		for (let i = start; i >= 0 && i < parent.childCount; i += dir) {
			const node = parent.child(i);
			if (nodeControlOf(node)?.value) return true;
			if (node.isTextblock) break;
		}
	}
	return false;
}

/** Text typed inside a single mark leaves both of its boundaries intact. */
function textEditChangesControls(doc: ProseMirrorNode, step: ReplaceStep): boolean | undefined {
	const { from, to, slice } = step;
	if (slice.openStart || slice.openEnd) return undefined;
	const $from = doc.resolve(from);
	const $to = doc.resolve(to);
	if (!$from.parent.isTextblock || !$from.sameParent($to)) return undefined;
	const left = controlMark($from.nodeBefore);
	const right = controlMark($to.nodeAfter);
	let textOnly = true;
	let insertedControl = false;
	let extendsControl = from === to && !!left && !!right && left.eq(right);
	slice.content.forEach((node) => {
		textOnly &&= node.isText;
		const mark = controlMark(node);
		insertedControl ||= !!mark;
		extendsControl &&= !!left && !!mark && mark.eq(left);
	});
	if (!textOnly) return undefined;
	if (extendsControl) return false;
	const type = doc.type.schema.marks.contentControl;
	return insertedControl || (!!type && doc.rangeHasMark(from, to, type)) || !!(left && right);
}

function replacementChangesControls(
	doc: ProseMirrorNode,
	step: ReplaceStep | ReplaceAroundStep,
	hasTags: boolean,
): boolean {
	if (hasTags && step instanceof ReplaceStep) {
		// A whole-document replacement, such as Yjs sync, invalidates the existing positions.
		if (step.from === 0 && step.to === doc.content.size) return true;
		const textEdit = textEditChangesControls(doc, step);
		if (textEdit !== undefined) return textEdit;
	}

	let found = false;
	step.slice.content.descendants((node) => {
		found ||= holdsControl(node, !hasTags);
		return !found;
	});
	if (found || !hasTags) return found;

	const { from, to } = step;
	doc.nodesBetween(from, to, (node) => {
		found ||= holdsControl(node, false);
		return !found;
	});
	if (found) return true;
	const $from = doc.resolve(from);
	const $to = doc.resolve(to);
	// An inserted inline atom can separate two marked runs without replacing either.
	return (
		!!controlMark($from.nodeBefore) ||
		!!controlMark($to.nodeAfter) ||
		besideBlockControl($from, -1) ||
		besideBlockControl($to, 1)
	);
}

/** Whether mapping tag positions alone can miss an added, removed or regrouped control. */
export function changesContentControlTags(tr: Transaction, hasTags: boolean): boolean {
	return tr.steps.some((step, index) => {
		const before = tr.docs[index];
		if (!before) return false;
		if (
			step instanceof AddMarkStep ||
			step instanceof RemoveMarkStep ||
			step instanceof AddNodeMarkStep ||
			step instanceof RemoveNodeMarkStep
		) {
			return (
				step.mark.type.name === 'contentControl' && (hasTags || isTagged(controlOfMark(step.mark)))
			);
		}
		if (step instanceof AttrStep) {
			const after = tr.docs[index + 1] ?? tr.doc;
			const oldNode = before.nodeAt(step.pos);
			const newNode = after.nodeAt(step.pos);
			if (!oldNode || !newNode) return false;
			const previous = nodeControlOf(oldNode)?.value;
			const next = nodeControlOf(newNode)?.value;
			return previous !== next && (hasTags || isTagged(next));
		}
		return (
			(step instanceof ReplaceStep || step instanceof ReplaceAroundStep) &&
			replacementChangesControls(before, step, hasTags)
		);
	});
}
