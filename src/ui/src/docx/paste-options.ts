/**
 * Word's Paste Options: after a paste the editor remembers the pasted range and what the
 * clipboard held, so Keep Source Formatting, Merge Formatting and Keep Text Only can paste it
 * again in place, as one undoable change. Typing, or Escape, ends the offer, as in Word.
 */
import type { WordPasteMode } from 'ooxml-core/docx';
import { closeHistory } from 'prosemirror-history';
import { Fragment, Slice, type Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin, PluginKey, TextSelection, type Transaction } from 'prosemirror-state';
import { ReplaceStep } from 'prosemirror-transform';
import type { EditorView } from 'prosemirror-view';

export type PasteOption = WordPasteMode | 'text';

/** What the clipboard held for one paste. */
export interface PastePayload {
	html: string | null;
	text: string;
}

export interface PastedRange {
	from: number;
	to: number;
	payload: PastePayload;
	option: PasteOption;
	/** The style of the paragraph the text was first pasted into, which Merge Formatting takes. */
	targetStyle: string;
}

/** Shared between the plugin's state and the commands that paste again. */
export class PasteController {
	/** The payload the next paste transaction comes from, captured before ProseMirror reads it. */
	payload: PastePayload | null = null;
	/** How the next paste is applied; `source` for an ordinary paste. */
	option: PasteOption = 'source';
	/** Set when pasting again: the first paste's destination style. */
	targetStyle: string | null = null;
}

export const pasteOptionsKey = new PluginKey<PastedRange | null>('dvePasteOptions');
const DISMISS = 'dismiss';

/** The marks Merge Formatting keeps: emphasis and links, never fonts, colours or styles. */
const MERGE_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'verticalAlign', 'link']);

function mergeFragment(fragment: Fragment, style: string): Fragment {
	const nodes: ProseMirrorNode[] = [];
	fragment.forEach((node) => {
		if (node.isText) {
			nodes.push(node.mark(node.marks.filter((mark) => MERGE_MARKS.has(mark.type.name))));
			return;
		}
		const content = mergeFragment(node.content, style);
		if (node.type.name !== 'paragraph') {
			nodes.push(node.copy(content));
			return;
		}
		// The text takes the paragraph formatting of where it lands; lists stay lists.
		const attrs = node.type.create().attrs;
		nodes.push(
			node.type.create(
				{ ...attrs, style, numId: node.attrs.numId, ilvl: node.attrs.ilvl },
				content,
				node.marks,
			),
		);
	});
	return Fragment.from(nodes);
}

/** A pasted slice as Merge Formatting pastes it into a paragraph styled `style`. */
export function mergeSlice(slice: Slice, style: string): Slice {
	return new Slice(mergeFragment(slice.content, style), slice.openStart, slice.openEnd);
}

/** Whether a transaction is an edit that ends the offer (typing, deleting, another paste). */
function endsOffer(tr: Transaction): boolean {
	if (tr.getMeta(pasteOptionsKey) === DISMISS) return true;
	if (tr.getMeta('appendedTransaction')) return false;
	// Attribute repairs (paragraph ids, note numbers) only move the range.
	return tr.steps.some((step) => step instanceof ReplaceStep);
}

/** Tracks the last paste; `controller` is filled by the paste event and the clipboard commands. */
export function pasteOptionsPlugin(controller: PasteController): Plugin<PastedRange | null> {
	return new Plugin<PastedRange | null>({
		key: pasteOptionsKey,
		controller,
		state: {
			init: () => null,
			apply(tr, value, oldState, newState) {
				if (tr.getMeta('uiEvent') === 'paste' && tr.docChanged && controller.payload) {
					const range: PastedRange = {
						from: tr.mapping.map(oldState.selection.from, -1),
						to: newState.selection.to,
						payload: controller.payload,
						option: controller.option,
						targetStyle:
							controller.targetStyle ?? String(oldState.selection.$from.parent.attrs.style ?? ''),
					};
					controller.payload = null;
					controller.option = 'source';
					controller.targetStyle = null;
					return range;
				}
				if (!value) return null;
				if (endsOffer(tr)) return null;
				if (!tr.docChanged) return value;
				const from = tr.mapping.map(value.from, -1);
				const to = tr.mapping.map(value.to, 1);
				return to > from ? { ...value, from, to } : null;
			},
		},
		props: {
			handleDOMEvents: {
				paste(_view, event) {
					const data = event.clipboardData;
					controller.payload = data
						? { html: data.getData('text/html') || null, text: data.getData('text/plain') }
						: null;
					return false;
				},
			},
			transformPasted(slice, view) {
				const { $from } = view.state.selection;
				let result = slice;
				if (controller.option !== 'source') {
					const style = controller.targetStyle ?? $from.parent.attrs.style;
					result = mergeSlice(result, String(style ?? ''));
				}
				// Pasting again over a range that began a paragraph replaces that paragraph too, so
				// its first paragraph takes the new choice's formatting rather than the old one's.
				if (
					controller.targetStyle !== null &&
					result.openStart > 0 &&
					$from.parentOffset === 0 &&
					result.content.firstChild?.type.name === 'paragraph'
				)
					result = new Slice(result.content, 0, result.openEnd);
				return result;
			},
		},
	});
}

const controllerOf = (view: EditorView): PasteController | undefined =>
	(pasteOptionsKey.get(view.state)?.spec as { controller?: PasteController } | undefined)
		?.controller;

/** Pastes `payload` as `option` at the selection, recording it for Paste Options. */
export function pastePayload(
	view: EditorView,
	payload: PastePayload,
	option: PasteOption,
	targetStyle: string | null = null,
): boolean {
	if (!view.editable) return false;
	const controller = controllerOf(view);
	if (option !== 'text' && !payload.html) return false;
	if (controller) {
		controller.payload = payload;
		controller.option = option;
		controller.targetStyle = targetStyle;
	}
	view.focus();
	const pasted =
		option === 'text' ? view.pasteText(payload.text) : view.pasteHTML(payload.html ?? '');
	// A paste that changed nothing leaves the controller ready for the next real one.
	if (controller) {
		controller.payload = null;
		controller.option = 'source';
		controller.targetStyle = null;
	}
	return pasted;
}

/** The last paste, while its Paste Options are on offer. */
export const pastedRange = (view: EditorView): PastedRange | null =>
	pasteOptionsKey.getState(view.state) ?? null;

/** Pastes the last paste again as `option`, replacing it as one undoable change. */
export function applyPasteOption(view: EditorView, option: PasteOption): boolean {
	const range = pastedRange(view);
	if (!range || !view.editable || range.option === option) return false;
	if (option !== 'text' && !range.payload.html) return false;
	const { doc } = view.state;
	const selection = TextSelection.between(doc.resolve(range.from), doc.resolve(range.to));
	// The selection carries no change of its own; closing the group keeps the swap one step.
	view.dispatch(closeHistory(view.state.tr.setSelection(selection)));
	return pastePayload(view, range.payload, option, range.targetStyle);
}

/** Withdraws the offer (Escape). */
export function dismissPasteOptions(view: EditorView): boolean {
	if (!pastedRange(view)) return false;
	view.dispatch(view.state.tr.setMeta(pasteOptionsKey, DISMISS));
	return true;
}
