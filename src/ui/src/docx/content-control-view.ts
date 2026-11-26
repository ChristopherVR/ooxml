// How the editor shows and operates content controls: the active outline, a click on a check box,
// and the list and calendar popups of drop-down, combo box and date picker controls, inline or
// around a table cell.
import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import {
	chooseDate,
	chooseListItem,
	controlAtSelection,
	toggleCheckboxControl,
} from 'ooxml-core/docx/ui';
import { canEditControlContent, controlDate, type ContentControl } from 'ooxml-core/docx';
import { openCalendarPopover } from './content-control-calendar';
import { contentControlTagWidgets } from './content-control-tag-widgets';
import { openListPopover } from './ribbon-popover';

/** The control a click landed on, read back from the span the mark drew. */
function controlOfElement(
	target: EventTarget | null,
): { control: ContentControl; el: HTMLElement } | null {
	const el = (target as HTMLElement | null)?.closest?.<HTMLElement>('[data-sdt]');
	if (!el?.dataset.sdt) return null;
	try {
		return { control: JSON.parse(el.dataset.sdt) as ContentControl, el };
	} catch {
		return null;
	}
}

/**
 * Where the control drawn as `el` starts in the document: controls can share an id (a file from
 * another tool may repeat one), and the commands act on the control around this position.
 */
function positionOf(view: EditorView, el: HTMLElement): number | undefined {
	if (!el.isConnected) return undefined;
	try {
		return view.posAtDOM(el, 0);
	} catch {
		return undefined;
	}
}

/** Opens the popup the control's kind calls for. Returns whether a popup (or toggle) took the click. */
function operate(
	view: EditorView,
	control: ContentControl,
	el: HTMLElement,
	event: MouseEvent,
): boolean {
	if (!canEditControlContent(control)) return control.kind === 'checkbox';
	if (control.kind === 'checkbox') {
		const tr = toggleCheckboxControl(view.state, control.id, positionOf(view, el));
		if (tr) view.dispatch(tr);
		return true;
	}
	const box = el.getBoundingClientRect();
	const onArrow = event.clientX > box.right - 24;
	if (
		(control.kind === 'dropDownList' || (control.kind === 'comboBox' && onArrow)) &&
		control.items?.length
	) {
		const items = control.items.map((item) => item.display);
		openListPopover(el, items, el.textContent ?? '', false, (display) => {
			const index = items.indexOf(display);
			const tr =
				index >= 0
					? chooseListItem(view.state, control.id, index, positionOf(view, el))
					: undefined;
			if (tr) view.dispatch(tr);
			view.focus();
		});
		return true;
	}
	if (control.kind === 'date') {
		openCalendarPopover(el, controlDate(control), control.dateLocale ?? 'en-US', (date) => {
			const tr = chooseDate(view.state, control.id, date, positionOf(view, el));
			if (tr) view.dispatch(tr);
			view.focus();
		});
		return true;
	}
	return false;
}

export function contentControlViewPlugin(): Plugin {
	const plugin: Plugin<DecorationSet> = new Plugin({
		state: {
			init: (_config, state) => contentControlTagWidgets(state.doc),
			apply: (tr, tags) => (tr.docChanged ? contentControlTagWidgets(tr.doc) : tags),
		},
		props: {
			decorations(state) {
				const tags = plugin.getState(state) ?? DecorationSet.empty;
				const active = controlAtSelection(state);
				if (!active) return tags;
				if (active.level === 'inline')
					return tags.add(state.doc, [
						Decoration.inline(active.from, active.to, { class: 'dve-sdt-active' }),
					]);
				if (active.level === 'cell')
					return tags.add(state.doc, [
						Decoration.node(active.from, active.to, { class: 'dve-sdt-active' }),
					]);
				const decorations: Decoration[] = [];
				state.doc.nodesBetween(active.from, active.to, (node, pos) => {
					if (node.type.name === 'paragraph' && node.attrs.contentControl?.id === active.control.id)
						decorations.push(
							Decoration.node(pos, pos + node.nodeSize, { class: 'dve-sdt-active' }),
						);
					return node.type.name !== 'paragraph';
				});
				return tags.add(state.doc, decorations);
			},
			handleClick(view, _pos, event) {
				if (!view.editable) return false;
				const found = controlOfElement(event.target);
				return found ? operate(view, found.control, found.el, event) : false;
			},
		},
	});
	return plugin;
}
