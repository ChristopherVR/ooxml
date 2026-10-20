// How the editor operates legacy form fields: a click on a check box toggles it, a click on a
// drop-down opens its entries, and Word's form field shading follows the document's setting.
import type { DocumentModel } from 'ooxml-core/docx';
import { formFieldOfNode, selectFormDropDownAt, toggleFormCheckboxAt } from 'ooxml-core/docx/ui';
import { Plugin } from 'prosemirror-state';
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
	if (!field || field.enabled === false) return Boolean(field);
	if (field.kind === 'checkbox') {
		const tr = toggleFormCheckboxAt(view.state, pos);
		if (tr) view.dispatch(tr);
		return true;
	}
	const entries = field.dropDown?.entries ?? [];
	if (field.kind !== 'dropDown' || !entries.length) return false;
	openListPopover(el, entries, el.textContent ?? '', false, (entry) => {
		const tr = selectFormDropDownAt(view.state, pos, entries.indexOf(entry));
		if (tr) view.dispatch(tr);
		view.focus();
	});
	return true;
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
		},
	});
}
