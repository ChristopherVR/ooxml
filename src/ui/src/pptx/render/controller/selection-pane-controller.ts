/**
 * The Selection Pane as one state owner for every binding: the row being renamed with its draft
 * and the label it was seeded with, and the row being dragged with the row under it; plus the one
 * row label every binding draws.
 *
 * Each binding kept its own copy of this pane, and they disagreed: React labelled an unnamed
 * shape in English whatever the locale ("Shape 3"), Vue translated it, vanilla capitalised the
 * type id ("Smartart 2"), Angular showed the element id and Svelte the lower-case type numbered
 * from the top; the rename box started from that label in three bindings, from the id in Angular
 * and from the bare name in Svelte; and vanilla let a read-only deck be reordered by dragging.
 * Now the label is {@link selectionPaneLabel} (name, else the text, else the translated
 * "{type} {number}"), a rename starts from that label and only an edited commit writes
 * (`resolveSelectionPaneRename`), and renaming and dragging are refused on a read-only deck.
 *
 * Selecting a row and toggling its visibility stay the binding's own commands; the controller
 * reports a rename and a reorder through its deps.
 *
 * @module render/controller/selection-pane-controller
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { resolveSelectionPaneRename } from '../selection-pane-rename';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/** What a row is called: its name, its text, or its type and position. */
export type SelectionPaneLabel =
	| { readonly kind: 'name'; readonly text: string }
	| { readonly kind: 'text'; readonly text: string }
	| { readonly kind: 'type'; readonly typeKey: string; readonly number: number };

/** The dictionary key of each element type's label. */
export const SELECTION_PANE_TYPE_KEYS: Readonly<Record<string, string>> = Object.freeze({
	text: 'pptx.elementType.textBox',
	shape: 'pptx.elementType.shape',
	connector: 'pptx.elementType.connector',
	image: 'pptx.elementType.image',
	picture: 'pptx.elementType.picture',
	chart: 'pptx.elementType.chart',
	table: 'pptx.elementType.table',
	smartArt: 'pptx.elementType.smartArt',
	media: 'pptx.elementType.media',
	group: 'pptx.elementType.group',
	ink: 'pptx.elementType.ink',
	ole: 'pptx.elementType.object',
	unknown: 'pptx.elementType.object',
});

/** The label of the element at stacking position `index` (0 is the back-most). */
export function selectionPaneLabel(element: PptxElement, index: number): SelectionPaneLabel {
	const name = element.name?.trim();
	if (name) {
		return { kind: 'name', text: name };
	}
	const text = hasTextProperties(element) ? element.text?.trim() : undefined;
	if (text) {
		return { kind: 'text', text: text.slice(0, 32) };
	}
	return {
		kind: 'type',
		typeKey: SELECTION_PANE_TYPE_KEYS[element.type] ?? 'pptx.elementType.object',
		number: index + 1,
	};
}

/** The label as text, translated with the binding's translator. */
export function formatSelectionPaneLabel(
	label: SelectionPaneLabel,
	t: (key: string, params?: Record<string, string | number>) => string,
): string {
	return label.kind === 'type'
		? t('pptx.selectionPane.elementLabel', { type: t(label.typeKey), number: label.number })
		: label.text;
}

/** One pane row: the element and its stacking position. */
export interface SelectionPaneRow {
	readonly element: PptxElement;
	/** Position in the slide's element list (0 is the back-most). */
	readonly index: number;
	readonly label: SelectionPaneLabel;
}

/** The pane's rows, top-most first (PowerPoint's order). */
export function selectionPaneRows(elements: readonly PptxElement[]): SelectionPaneRow[] {
	return elements
		.map((element, index) => ({ element, index, label: selectionPaneLabel(element, index) }))
		.reverse();
}

export interface SelectionPaneState {
	/** The element being renamed, or null. */
	readonly renamingId: string | null;
	/** The rename box's text. */
	readonly renameDraft: string;
	/** The label the box was seeded with: an unedited commit writes nothing. */
	readonly renameSeed: string;
	/** Stacking position of the row being dragged, or null. */
	readonly dragIndex: number | null;
	/** Stacking position of the row under it, or null. */
	readonly dragOverIndex: number | null;
}

export interface SelectionPaneControllerDeps {
	canEdit(): boolean;
	/** Write `name` (`''` clears it) on element `id`. */
	onRename(id: string, name: string): void;
	/** Move the element at stacking position `from` to `to`. */
	onReorder(from: number, to: number): void;
}

export interface SelectionPaneController extends PptxController<SelectionPaneState> {
	/** Open the rename box on `id`, seeded with the row's label text. */
	startRename(id: string, label: string): boolean;
	setRenameDraft(text: string): void;
	/** Enter or blur: write the name when it was edited; returns whether it wrote. */
	commitRename(): boolean;
	/** Escape. */
	cancelRename(): void;
	dragStart(index: number): boolean;
	dragOver(index: number): void;
	drop(index: number): void;
	dragEnd(): void;
}

const IDLE: SelectionPaneState = Object.freeze({
	renamingId: null,
	renameDraft: '',
	renameSeed: '',
	dragIndex: null,
	dragOverIndex: null,
});

export function createSelectionPaneController(
	deps: SelectionPaneControllerDeps,
): SelectionPaneController {
	const store = createControllerStore<SelectionPaneState>(IDLE);
	const patch = (next: Partial<SelectionPaneState>): void =>
		store.update((current) => ({ ...current, ...next }));
	const endDrag = (): void => {
		const { dragIndex, dragOverIndex } = store.getSnapshot();
		if (dragIndex !== null || dragOverIndex !== null) {
			patch({ dragIndex: null, dragOverIndex: null });
		}
	};
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		startRename(id, label) {
			if (!deps.canEdit()) {
				return false;
			}
			patch({ renamingId: id, renameDraft: label, renameSeed: label });
			return true;
		},
		setRenameDraft(text) {
			if (store.getSnapshot().renamingId !== null) {
				patch({ renameDraft: text });
			}
		},
		commitRename() {
			const { renamingId, renameSeed, renameDraft } = store.getSnapshot();
			if (renamingId === null) {
				return false;
			}
			// Closed first: the commit's own blur re-enters here and must be a no-op.
			patch({ renamingId: null });
			const commit = resolveSelectionPaneRename(renameSeed, renameDraft);
			if (!commit || !deps.canEdit()) {
				return false;
			}
			deps.onRename(renamingId, commit.name);
			return true;
		},
		cancelRename() {
			if (store.getSnapshot().renamingId !== null) {
				patch({ renamingId: null });
			}
		},
		dragStart(index) {
			if (!deps.canEdit()) {
				return false;
			}
			patch({ dragIndex: index, dragOverIndex: null });
			return true;
		},
		dragOver(index) {
			const { dragIndex, dragOverIndex } = store.getSnapshot();
			if (dragIndex !== null && dragIndex !== index && dragOverIndex !== index) {
				patch({ dragOverIndex: index });
			}
		},
		drop(index) {
			const { dragIndex } = store.getSnapshot();
			endDrag();
			if (dragIndex !== null && dragIndex !== index && deps.canEdit()) {
				deps.onReorder(dragIndex, index);
			}
		},
		dragEnd: endDrag,
	};
}
