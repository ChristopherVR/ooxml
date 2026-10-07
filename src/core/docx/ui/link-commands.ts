import type { Mark } from 'prosemirror-model';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { TextRun } from '../model';
import { inlineNodeRun } from './run-adapter';
import { hasInlineRunAttributes, setInlineRunFormatting } from './inline-formatting';
import { isNavigableHref } from './inline-content-schema';

export type LinkTarget = NonNullable<TextRun['link']>;

/** A contiguous link can span ordinary text and attribute-backed inline elements. */
export function wordLinkRangeAt(
	state: EditorState,
	pos: number,
): { from: number; to: number; mark: Mark } | undefined {
	const type = state.schema.marks.link;
	if (!type) return;
	const $pos = state.doc.resolve(pos);
	const items: { from: number; to: number; link?: Mark }[] = [];
	let cursor = $pos.start();
	$pos.parent.forEach((child) => {
		const target = inlineNodeRun(child)?.link;
		const link = target ? type.create(target) : undefined;
		items.push({ from: cursor, to: cursor + child.nodeSize, ...(link && { link }) });
		cursor += child.nodeSize;
	});
	let index = items.findIndex((item) => item.link && item.from <= pos && pos < item.to);
	if (index < 0) index = items.findIndex((item) => item.link && item.from < pos && pos <= item.to);
	if (index < 0) return;
	const mark = items[index]?.link;
	if (!mark) return;
	let start = index;
	let end = index;
	while (start > 0) {
		if (!items[start - 1]?.link?.eq(mark)) break;
		start--;
	}
	while (end < items.length - 1) {
		if (!items[end + 1]?.link?.eq(mark)) break;
		end++;
	}
	return { from: items[start]!.from, to: items[end]!.to, mark };
}

function setRangeLink(tr: Transaction, from: number, to: number, target?: LinkTarget): void {
	const type = tr.doc.type.schema.marks.link!;
	tr.removeMark(from, to, type);
	tr.doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isInline) return;
		if (hasInlineRunAttributes(node) && node.type.name !== 'equation') {
			const run = inlineNodeRun(node)!;
			delete run.link;
			if (target) run.link = { ...target };
			setInlineRunFormatting(tr, pos, run);
		} else if (target)
			tr.addMark(Math.max(from, pos), Math.min(to, pos + node.nodeSize), type.create(target));
	});
}

/** Retargets a selection or existing link, or inserts linked text at an empty cursor. */
export function applyWordLink(target: LinkTarget, displayText?: string): Command {
	return (state, dispatch, view) => {
		if (view?.editable === false || !state.schema.marks.link || (!target.href && !target.anchor))
			return false;
		if (target.href && !isNavigableHref(target.href))
			throw new Error('Links must start with http://, https:// or mailto:.');
		if (!dispatch) return true;
		const { from, to, empty } = state.selection;
		const tr = state.tr;
		const existing = empty ? wordLinkRangeAt(state, from) : { from, to };
		if (existing) setRangeLink(tr, existing.from, existing.to, target);
		else {
			const type = state.schema.marks.link;
			const marks = (state.storedMarks ?? state.selection.$from.marks()).filter(
				(mark) => mark.type !== type,
			);
			tr.insert(
				from,
				state.schema.text(displayText || target.href || target.anchor!, [
					...marks,
					type.create(target),
				]),
			);
		}
		dispatch(tr.scrollIntoView());
		return true;
	};
}

export const removeWordLink: Command = (state, dispatch, view) => {
	if (view?.editable === false || !state.schema.marks.link) return false;
	const { from, to, empty } = state.selection;
	const range = empty ? wordLinkRangeAt(state, from) : { from, to };
	if (!range) return false;
	if (dispatch) {
		const tr = state.tr;
		setRangeLink(tr, range.from, range.to);
		dispatch(tr);
	}
	return true;
};
