import type { Command } from 'prosemirror-state';
import type { Node, Mark } from 'prosemirror-model';
import type { TextRun } from '../model';
import { DIRECT_RUN_PROPERTY_KEYS } from '../run-formatting';
import { inlineNodeRun } from './run-adapter';
import { marksForRun } from './run-marks';
import {
	RUN_FORMAT_MARKS,
	hasInlineRunAttributes,
	setInlineRunFormatting,
} from './inline-formatting';

export type RunFormattingPatch = {
	[K in (typeof DIRECT_RUN_PROPERTY_KEYS)[number]]?: TextRun[K] | undefined;
};

function patchedRun(node: Node, patch: RunFormattingPatch): TextRun | undefined {
	const run = inlineNodeRun(node);
	if (!run) return undefined;
	const properties = run as unknown as Record<string, unknown>;
	for (const [key, value] of Object.entries(patch)) {
		if (value === undefined) delete properties[key];
		else properties[key] = value;
	}
	return run;
}

const formattingMarks = (node: Node, run: TextRun): Mark[] =>
	marksForRun(run, node.type.schema).filter((mark) => RUN_FORMAT_MARKS.has(mark.type.name));

/** Applies direct formatting to text and supported inline atoms, retaining independent history. */
export function applyRunFormattingPatch(patch: RunFormattingPatch): Command {
	return (state, dispatch, view) => {
		if (view && !view.editable) return false;
		if (!dispatch) return true;
		const tr = state.tr;
		const { selection } = state;
		if (selection.empty) {
			const marks = state.storedMarks ?? selection.$from.marks();
			const node = state.schema.text('x', marks);
			const run = patchedRun(node, patch)!;
			dispatch(
				tr.setStoredMarks([
					...marks.filter((mark) => !RUN_FORMAT_MARKS.has(mark.type.name)),
					...formattingMarks(node, run),
				]),
			);
			return true;
		}
		state.doc.nodesBetween(selection.from, selection.to, (node, pos) => {
			if (!node.isInline || node.type.name === 'equation') return;
			const run = patchedRun(node, patch);
			if (!run) return;
			if (hasInlineRunAttributes(node)) {
				setInlineRunFormatting(tr, pos, run);
				return;
			}
			const from = Math.max(pos, selection.from);
			const to = Math.min(pos + node.nodeSize, selection.to);
			for (const mark of node.marks)
				if (RUN_FORMAT_MARKS.has(mark.type.name)) tr.removeMark(from, to, mark);
			for (const mark of formattingMarks(node, run)) tr.addMark(from, to, mark);
		});
		if (tr.docChanged) dispatch(tr);
		return true;
	};
}

/** Clears modeled direct run properties while preserving links, fields, comments and revisions. */
export const clearDirectRunFormatting = applyRunFormattingPatch(
	Object.fromEntries(DIRECT_RUN_PROPERTY_KEYS.map((key) => [key, undefined])) as RunFormattingPatch,
);
