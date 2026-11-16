import {
	createSelectionPaneController,
	formatSelectionPaneLabel,
	restoreEditorKeyboardFocus,
	selectionPaneLabel,
	selectionPaneRows,
} from 'ooxml-ui/pptx';
import type { SelectionPaneController, SelectionPaneRow } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { Translator } from '../i18n';
import { createEl } from '../render';

/** What the pane lists: the active slide's objects and the current selection. */
export interface SelectionPaneModel {
	elements: readonly PptxElement[];
	selectedIds: readonly string[];
}

export interface SelectionPaneOptions extends SelectionPaneModel {
	onSelect(id: string): void;
	onToggleHidden(id: string): void;
	onReorder(from: number, to: number): void;
	/** Whether the host allows edits: renaming and dragging are refused otherwise. */
	canEdit?(): boolean;
	/** Rename commit: the trimmed name, or `''` to clear it. */
	onRename(id: string, name: string): void;
	/**
	 * Live model feed. Without it the pane is a snapshot of the deck as it was
	 * when it opened, which is how an undone rename kept showing the new name:
	 * the model went back, the rows did not. Returns an unsubscribe function.
	 */
	subscribe?(listener: (model: SelectionPaneModel) => void): () => void;
	/** Invoked when the pane's own close button tears it down. */
	onClose?(): void;
}

/** Handle so the caller can release the pane (and its subscription). */
export interface SelectionPaneHandle {
	close(): void;
}

export function openSelectionPane(
	doc: Document,
	host: HTMLElement,
	t: Translator,
	options: SelectionPaneOptions,
): SelectionPaneHandle {
	host.querySelector('[data-pptx-selection-pane]')?.remove();
	const pane = createEl(doc, 'aside', 'pptxv-workspace-pane');
	pane.dataset.pptxSelectionPane = 'true';
	const header = createEl(doc, 'header');
	const title = createEl(doc, 'h2');
	title.textContent = t('pptx.selectionPane.title');
	const close = createEl(doc, 'button');
	close.type = 'button';
	close.textContent = '×';
	close.setAttribute('aria-label', t('pptx.selectionPane.close'));
	header.append(title, close);
	pane.appendChild(header);

	const list = createEl(doc, 'div', 'pptxv-workspace-list');
	// The shared rename box and drag state; the deck changes go out through the options.
	const controller = createSelectionPaneController({
		canEdit: () => options.canEdit?.() ?? true,
		onRename: options.onRename,
		onReorder: options.onReorder,
	});
	let model: SelectionPaneModel = { elements: options.elements, selectedIds: options.selectedIds };
	let renderedElements: readonly PptxElement[] | null = null;
	const rowsById = new Map<string, HTMLElement>();

	/** Repaint the selection highlight without touching the row nodes. */
	const paintSelection = (): void => {
		for (const [id, row] of rowsById) {
			row.classList.toggle('is-selected', model.selectedIds.includes(id));
		}
	};
	const renderRows = (): void => {
		// Selection-only changes must NOT rebuild: a row whose node is swapped
		// between the two halves of a double-click never sees the `dblclick` (it
		// lands on the common ancestor instead), which would make rename
		// unreachable, since the first click of the pair selects the row.
		if (renderedElements === model.elements) {
			paintSelection();
			return;
		}
		renderedElements = model.elements;
		rowsById.clear();
		list.replaceChildren();
		if (!model.elements.length) {
			const empty = createEl(doc, 'p');
			empty.textContent = t('pptx.selectionPane.empty');
			list.appendChild(empty);
			return;
		}
		for (const paneRow of selectionPaneRows(model.elements)) {
			const row = buildRow(doc, t, paneRow, options, controller, model.selectedIds);
			rowsById.set(paneRow.element.id, row);
			list.appendChild(row);
		}
	};
	renderRows();
	pane.appendChild(list);

	const unsubscribe =
		options.subscribe?.((next) => {
			model = next;
			// A rename in flight owns the DOM: rebuilding under it would drop the
			// input mid-keystroke. The commit re-renders as soon as it is done.
			if (list.querySelector('.pptxv-selection-rename')) {
				return;
			}
			renderRows();
		}) ?? null;
	const handle: SelectionPaneHandle = {
		close() {
			unsubscribe?.();
			pane.remove();
		},
	};
	close.addEventListener('click', () => {
		handle.close();
		options.onClose?.();
	});
	host.appendChild(pane);
	return handle;
}

/** One object row: drag handle, name label (double-click renames), eye toggle. */
function buildRow(
	doc: Document,
	t: Translator,
	{ element, index, label }: SelectionPaneRow,
	options: SelectionPaneOptions,
	controller: SelectionPaneController,
	selectedIds: readonly string[],
): HTMLElement {
	const row = createEl(doc, 'div', 'pptxv-selection-row');
	row.draggable = options.canEdit?.() ?? true;
	row.dataset.index = String(index);
	row.classList.toggle('is-selected', selectedIds.includes(element.id));
	const grip = createEl(doc, 'span');
	grip.textContent = '☰';
	const name = createEl(doc, 'button');
	name.type = 'button';
	// E2E contract: the row's name label, double-clicked to rename.
	name.dataset.pptxSelectionName = 'true';
	name.textContent = formatSelectionPaneLabel(label, t);
	name.addEventListener('click', () => options.onSelect(element.id));
	name.addEventListener('dblclick', () =>
		startRename(doc, t, row, name, element, index, controller),
	);
	const visible = createEl(doc, 'button');
	visible.type = 'button';
	visible.textContent = element.hidden ? '○' : '◉';
	visible.setAttribute(
		'aria-label',
		t(element.hidden ? 'pptx.selectionPane.show' : 'pptx.selectionPane.hide'),
	);
	visible.addEventListener('click', () => options.onToggleHidden(element.id));
	row.addEventListener('dragstart', (event) => {
		event.dataTransfer?.setData('text/plain', String(index));
		controller.dragStart(index);
	});
	row.addEventListener('dragover', (event) => {
		event.preventDefault();
		controller.dragOver(index);
	});
	row.addEventListener('dragend', () => controller.dragEnd());
	row.addEventListener('drop', (event) => {
		event.preventDefault();
		controller.drop(index);
	});
	row.append(grip, name, visible);
	return row;
}

/**
 * Swap the name button for a text input, matching every binding: Enter and blur commit, Escape
 * cancels. The input is seeded with the row's label, and an unedited commit changes nothing (a
 * "Shape 3" fallback label is never written as a real name); an emptied box clears the name.
 */
function startRename(
	doc: Document,
	t: Translator,
	row: HTMLElement,
	name: HTMLButtonElement,
	element: PptxElement,
	index: number,
	controller: SelectionPaneController,
): void {
	const seed = name.textContent ?? '';
	if (!controller.startRename(element.id, seed)) {
		return;
	}
	const input = doc.createElement('input');
	input.type = 'text';
	input.className = 'pptxv-selection-rename';
	input.setAttribute('aria-label', t('pptx.selectionPane.renameElement'));
	input.value = seed;
	name.hidden = true;
	row.insertBefore(input, name);
	input.focus();
	input.select();
	input.addEventListener('input', () => controller.setRenameDraft(input.value));

	// Enter commits and removes the input, which also fires blur: the controller's own guard
	// keeps that second path from committing twice.
	const finish = (commit: boolean): void => {
		if (controller.getSnapshot().renamingId !== element.id) {
			return;
		}
		// The box closes first (the blur that focusing the root causes must find it closed), then
		// the keyboard goes back to the viewer root BEFORE the input leaves the document, or
		// focus falls outside the root this binding listens on and Ctrl+Z is dropped.
		let written: string | null = null;
		if (commit) {
			// The field's value as it stands (an IME or a programmatic fill may not fire `input`).
			controller.setRenameDraft(input.value);
			written = controller.commitRename() ? input.value.trim() : null;
		} else {
			controller.cancelRename();
		}
		restoreEditorKeyboardFocus(input);
		input.remove();
		name.hidden = false;
		if (written !== null) {
			// The live feed repaints the row; until then it shows the new name (or the label).
			name.textContent =
				written || formatSelectionPaneLabel(selectionPaneLabel({ ...element, name: '' }, index), t);
		}
	};
	input.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			finish(true);
		} else if (event.key === 'Escape') {
			finish(false);
		}
	});
	input.addEventListener('blur', () => finish(true));
}
