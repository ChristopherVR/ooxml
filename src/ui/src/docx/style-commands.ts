import {
	editStyleCatalogs,
	paragraphStyleId,
	styleCatalogsOf,
	type DocumentModel,
	type ParagraphFormatting,
	type RunFormatting,
	type StyleEdit,
} from 'ooxml-core/docx';
import { marksForRun } from 'ooxml-core/docx/ui';
import { runOf } from './run-styles';
import { splitBlockAs } from 'prosemirror-commands';
import { closeHistory } from 'prosemirror-history';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { characterStyleAtSelection } from './character-style-picker';
import { tableBorderContext } from './table-border-commands';

const RUN_FIELDS = [
	'bold',
	'italic',
	'fontFamily',
	'fontSize',
	'color',
	'fontTheme',
	'colorTheme',
] as const;
const PARAGRAPH_FIELDS = [
	'align',
	'justification',
	'spacingBeforeTwips',
	'spacingAfterTwips',
	'lineSpacingTwips',
	'lineSpacingRule',
] as const;

function runAtSelection(state: EditorState) {
	const { $from, empty } = state.selection;
	const selected = empty ? null : $from.nodeAfter;
	return runOf(selected ?? state.schema.text('x', state.storedMarks ?? $from.marks()));
}

export function selectionStyle(view: EditorView, model: DocumentModel) {
	const { attrs } = view.state.selection.$from.parent;
	const paragraphId = paragraphStyleId(model, attrs.style || undefined);
	const characterId = characterStyleAtSelection(view);
	const run = runAtSelection(view.state);
	const { bold, italic, fontFamily, fontSize, color, fontTheme, colorTheme } = run ?? {};
	const runFormatting: RunFormatting = {
		...(bold !== undefined && { bold }),
		...(italic !== undefined && { italic }),
		...(fontFamily !== undefined && { fontFamily }),
		...(fontSize !== undefined && { fontSize }),
		...(color !== undefined && { color }),
		...(fontTheme !== undefined && { fontTheme }),
		...(colorTheme !== undefined && { colorTheme }),
	};
	const paragraphFormatting: ParagraphFormatting = {
		...(attrs.align != null && { align: attrs.align }),
		...(attrs.justification != null && { justification: attrs.justification }),
		...(attrs.spacingBeforeTwips != null && { spacingBeforeTwips: attrs.spacingBeforeTwips }),
		...(attrs.spacingAfterTwips != null && { spacingAfterTwips: attrs.spacingAfterTwips }),
		...(attrs.lineSpacingTwips != null && { lineSpacingTwips: attrs.lineSpacingTwips }),
		...(attrs.lineSpacingRule != null && { lineSpacingRule: attrs.lineSpacingRule }),
	};
	const table = tableBorderContext(view.state);
	const tableId: string = table
		? (view.state.doc.nodeAt(table.tablePos)?.attrs.tableStyleId ?? '')
		: '';
	return { runFormatting, paragraphFormatting, paragraphId, characterId, tableId };
}

export function commitStyleEdit(view: EditorView, model: DocumentModel, edit: StyleEdit): boolean {
	if (!view.editable || model.protection?.enforced) return false;
	const result = editStyleCatalogs(styleCatalogsOf(model), edit);
	if (!result) return false;
	const { id, catalogs, linkedId } = result;
	const { state } = view;
	const { from, to, empty } = state.selection;
	const tr = state.tr.setDocAttribute('styleCatalogs', catalogs);
	if (edit.id === undefined) {
		const sample = selectionStyle(view, model);
		const char =
			edit.kind === 'character'
				? id
				: edit.kind === 'linked' &&
					  !empty &&
					  state.selection.$from.sameParent(state.selection.$to) &&
					  (state.selection.$from.parentOffset !== 0 ||
							state.selection.$to.parentOffset !== state.selection.$to.parent.content.size)
					? linkedId
					: undefined;
		if (char) {
			const mark = state.schema.marks.characterStyle;
			if (!mark) return false;
			if (empty) tr.addStoredMark(mark.create({ id: char }));
			else {
				moveSelectionFormatting(tr, from, to, sample.runFormatting);
				tr.removeMark(from, to, mark);
				tr.addMark(from, to, mark.create({ id: char }));
			}
		} else if (edit.kind === 'table') {
			const table = tableBorderContext(state);
			if (table) tr.setNodeAttribute(table.tablePos, 'tableStyleId', id);
		} else if (edit.kind === 'paragraph' || edit.kind === 'linked') {
			state.doc.nodesBetween(from, to, (node, pos) => {
				if (node.type.name !== 'paragraph') return;
				tr.setNodeAttribute(pos, 'style', id);
				for (const key of PARAGRAPH_FIELDS)
					if (
						sample.paragraphFormatting[key] !== undefined &&
						node.attrs[key] === sample.paragraphFormatting[key]
					)
						tr.setNodeAttribute(pos, key, null);
				moveSelectionFormatting(tr, pos + 1, pos + node.nodeSize - 1, sample.runFormatting);
			});
		}
	}
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}

/** Moves the sampled direct formatting to the style while keeping other runs' overrides. */
function moveSelectionFormatting(tr: Transaction, from: number, to: number, sample: RunFormatting) {
	tr.doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isText) return;
		const run = runOf(node);
		if (!run) return;
		let changed = false;
		for (const key of RUN_FIELDS) {
			if (sample[key] === undefined || JSON.stringify(run[key]) !== JSON.stringify(sample[key]))
				continue;
			delete run[key];
			changed = true;
		}
		if (!changed) return;
		const start = Math.max(from, pos);
		const end = Math.min(to, pos + node.nodeSize);
		const formattingMark = (name: string) =>
			['bold', 'italic', 'font', 'runProperties'].includes(name);
		for (const mark of node.marks)
			if (formattingMark(mark.type.name)) tr.removeMark(start, end, mark);
		for (const mark of marksForRun(run, node.type.schema))
			if (formattingMark(mark.type.name)) tr.addMark(start, end, mark);
	});
}

/** At the end of a paragraph, Enter uses the following style defined by its style. */
export function splitWithFollowingStyle(getModel: () => DocumentModel): Command {
	return (state, dispatch, view) => {
		const { $from, empty } = state.selection;
		if (
			!empty ||
			$from.parent.type.name !== 'paragraph' ||
			$from.parentOffset !== $from.parent.content.size
		)
			return false;
		const model = getModel();
		const styles = model.paragraphStyles?.styles;
		const style = paragraphStyleId(model, $from.parent.attrs.style || undefined);
		const next = style && styles?.[style]?.next;
		if (!next || !Object.hasOwn(styles ?? {}, next) || $from.parent.attrs.numId != null)
			return false;
		return splitBlockAs((node) => ({ type: node.type, attrs: { style: next } }))(
			state,
			dispatch,
			view,
		);
	};
}
