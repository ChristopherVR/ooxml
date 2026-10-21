// How the editor operates legacy form fields: a click on a check box toggles it, a click on a
// drop-down opens its entries, keys do the same beside the field, and Word's form field shading
// follows the document's setting.
import { formDropDownIndex, type DocumentModel } from 'ooxml-core/docx';
import {
	formFieldBeside,
	formFieldLocked,
	formFieldOfNode,
	selectFormDropDownAt,
	toggleFormCheckboxAt,
} from 'ooxml-core/docx/ui';
import { splitBlock } from 'prosemirror-commands';
import { Plugin, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { openListPopover } from './ribbon-popover';

/** The document position of the begin marker a form field element was drawn for. */
function markerPosition(view: EditorView, el: HTMLElement): number | undefined {
	const near = view.posAtDOM(el, 0);
	return [near, near - 1].find((pos) => pos >= 0 && view.nodeDOM(pos) === el);
}

function operate(view: EditorView, el: HTMLElement): boolean {
	const pos = markerPosition(view, el);
	if (pos === undefined) return false;
	const field = formFieldOfNode(view.state.doc.nodeAt(pos));
	// A disabled field, or one a content control locks, takes the click without changing.
	if (!field || field.enabled === false || formFieldLocked(view.state, pos)) return Boolean(field);
	if (field.kind === 'checkbox') {
		const tr = toggleFormCheckboxAt(view.state, pos);
		if (tr) view.dispatch(tr);
		return true;
	}
	const entries = field.dropDown?.entries ?? [];
	if (field.kind !== 'dropDown' || !entries.length) return false;
	// Entries can repeat a name, so the list goes by the entry's place, as `w:result` does. The
	// document can change while the list is open (a collaborator's edit), so the pick goes to the
	// field the list was opened for, wherever it is now, and to none once it is gone or redrawn.
	openListPopover(el, entries, formDropDownIndex(field) ?? -1, false, (_entry, index) => {
		const at = el.isConnected ? markerPosition(view, el) : undefined;
		const tr = at === undefined ? undefined : selectFormDropDownAt(view.state, at, index);
		if (tr) view.dispatch(tr);
		view.focus();
	});
	return true;
}

/**
 * Whether the key has an edit to make at the caret that no guard refuses: Space types a space,
 * Enter splits the paragraph.
 */
function keyEdits(view: EditorView, key: string): boolean {
	let tr: Transaction | undefined;
	if (key === ' ') tr = view.state.tr.insertText(' ');
	else splitBlock(view.state, (split) => (tr = split));
	return (
		Boolean(tr) &&
		view.state.plugins.every(
			(plugin) => plugin.spec.filterTransaction?.call(plugin, tr!, view.state) ?? true,
		)
	);
}

/** Where the browser's caret is in the editor's text, if it is there and maps to a position. */
function browserCaret(view: EditorView): number | undefined {
	const root = view.root as { getSelection?: () => Selection | null };
	const selection = root.getSelection?.() ?? document.getSelection();
	const node = selection?.focusNode;
	if (!node || !view.dom.contains(node)) return undefined;
	try {
		return view.posAtDOM(node, selection!.focusOffset);
	} catch {
		return undefined;
	}
}

/**
 * Operates the check box or drop-down beside the caret from the keyboard. Space or Enter toggles
 * a check box or opens a drop-down's entries where the key has nothing to type, as in a form
 * protected for filling in (Word moves to a field and Space toggles it); Alt+Down opens the
 * entries anywhere.
 */
function operateFromKey(view: EditorView, event: KeyboardEvent): boolean {
	if (event.ctrlKey || event.metaKey || event.shiftKey) return false;
	const list = event.altKey && event.key === 'ArrowDown';
	const press = !event.altKey && (event.key === ' ' || event.key === 'Enter');
	if ((!list && !press) || !view.state.selection.empty) return false;
	const pos = formFieldBeside(view.state.doc, view.state.selection.from);
	if (pos === undefined) return false;
	if (list && formFieldOfNode(view.state.doc.nodeAt(pos))?.kind !== 'dropDown') return false;
	// Right after a click the browser's caret can be where the editor has not read it yet: act
	// only on a field the caret the user sees is beside too.
	const caret = browserCaret(view);
	if (caret !== undefined && formFieldBeside(view.state.doc, caret) !== pos) return false;
	if (press && keyEdits(view, event.key)) return false;
	const el = view.nodeDOM(pos);
	return el instanceof HTMLElement && operate(view, el);
}

export function formFieldViewPlugin(model: () => DocumentModel): Plugin {
	return new Plugin({
		props: {
			attributes: () => ({
				'data-form-shading': model().formFieldShading === false ? 'off' : 'on',
			}),
			handleClick(view, _pos, event) {
				if (!view.editable) return false;
				const el = (event.target as HTMLElement | null)?.closest?.<HTMLElement>('.dve-form-field');
				return el ? operate(view, el) : false;
			},
			handleKeyDown: (view, event) => view.editable && operateFromKey(view, event),
		},
	});
}
