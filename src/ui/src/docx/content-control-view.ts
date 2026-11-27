// How the editor shows and operates content controls: the active outline, a click on a check box,
// and the list and calendar popups of drop-down, combo box and date picker controls, inline or
// around a table cell.
import { Plugin, TextSelection, type Transaction } from 'prosemirror-state';
import { AttrStep } from 'prosemirror-transform';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import {
	changesContentControlTags,
	chooseDate,
	chooseListItem,
	controlAtSelection,
	toggleCheckboxControl,
} from 'ooxml-core/docx/ui';
import { canEditControlContent, controlDate, type ContentControl } from 'ooxml-core/docx';
import { openCalendarPopover } from './content-control-calendar';
import { inlineCaretMarks, inlineCaretSide } from './content-control-tag-caret';
import {
	contentControlTagWidgets,
	widgetsWithCaretInside,
	type TagEdge,
} from './content-control-tag-widgets';
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

/** Paragraph attributes a tag's box depends on (`tagBoxStyle`), which mapping alone keeps stale. */
const changesTagBoxes = (tr: Transaction) =>
	tr.steps.some(
		(step) =>
			step instanceof AttrStep &&
			(step.attr === 'lineSpacingRule' || step.attr === 'lineSpacingTwips'),
	);

interface TagState {
	widgets: DecorationSet;
	inside: TagEdge | undefined;
}

export function contentControlViewPlugin(): Plugin {
	const plugin: Plugin<TagState> = new Plugin({
		state: {
			init: (_config, state) => {
				const widgets = contentControlTagWidgets(state.doc);
				return { widgets, inside: inlineCaretSide(widgets, state.selection) };
			},
			apply: (tr, tags, state) => {
				const { widgets: previous } = tags;
				const hasTags = previous !== DecorationSet.empty;
				const widgets =
					changesContentControlTags(tr, hasTags) || (hasTags && changesTagBoxes(tr))
						? contentControlTagWidgets(tr.doc)
						: tr.docChanged && hasTags
							? previous.map(tr.mapping, tr.doc)
							: previous;
				if (widgets === DecorationSet.empty && !hasTags) return tags;
				const inside = inlineCaretSide(
					widgets,
					tr.selection,
					{ tr, state, inside: tags.inside },
					tr.getMeta(plugin),
				);
				return { widgets, inside };
			},
		},
		appendTransaction: (_trs, _previous, state) => {
			const tags = plugin.getState(state);
			return tags && tags.widgets !== DecorationSet.empty
				? inlineCaretMarks(tags.widgets, tags.inside, state)
				: undefined;
		},
		props: {
			decorations(state) {
				const value = plugin.getState(state);
				const widgets = value?.widgets ?? DecorationSet.empty;
				const tags = value?.inside
					? widgetsWithCaretInside(state.doc, widgets, state.selection.from, value.inside)
					: widgets;
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
			handleDOMEvents: {
				mousedown(view, event) {
					if (event.button !== 0 || !(event.target instanceof HTMLElement)) return false;
					const tag = event.target.closest<HTMLElement>('.dve-sdt-tag');
					if (!tag) return false;
					event.preventDefault();
					const pos = view.posAtDOM(tag, 0);
					const tr = view.state.tr.setSelection(TextSelection.create(view.state.doc, pos));
					const edge = tag.dataset.tagEdge;
					if (tag.hasAttribute('data-inline-tag') && (edge === 'start' || edge === 'end'))
						tr.setMeta(plugin, edge);
					view.dispatch(tr);
					view.focus();
					return true;
				},
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
