import {
	isEditCancellation,
	visioSelectionIsOnPage,
	visioSizePositionCommand,
	visioSizePositionState,
} from 'ooxml-core/visio/ui';
import { canEditDiagram } from './controller';
import type { ViewerController, ViewerState } from './controller';

type InlineField = 'width' | 'height';

/**
 * The Width and Height boxes in the Size group of the Shape, Picture and Container Format tabs.
 * Each commits one resize of the selected shape through the same core command the Size & Position
 * pane uses, so the two always agree.
 */
export class ViewerInlineSize {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
	) {}
	wire(): () => void {
		const Abort = this.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		for (const input of this.inputs()) {
			input.addEventListener('change', () => void this.#commit(input), { signal: events.signal });
			input.addEventListener(
				'keydown',
				(event) => {
					if (event.key !== 'Enter' || event.isComposing) return;
					event.preventDefault();
					void this.#commit(input);
				},
				{ signal: events.signal },
			);
		}
		return () => events.abort();
	}
	inputs(): HTMLInputElement[] {
		return [...this.root.querySelectorAll<HTMLInputElement>('[data-inline-size]')];
	}
	#selection(state: ViewerState) {
		const page = state.document?.pages[state.pageIndex];
		const selected = state.selectedShapes.length === 1 ? state.selectedShape : null;
		return page && selected && visioSelectionIsOnPage(selected, page.id)
			? { page, id: selected.id }
			: undefined;
	}
	async #commit(input: HTMLInputElement): Promise<void> {
		const field = input.dataset.inlineSize as InlineField;
		const state = this.controller.state;
		const selection = this.#selection(state);
		if (!selection || state.loading || state.edit.busy || !canEditDiagram(state)) {
			this.render(state, true);
			return;
		}
		const command = Number.isFinite(input.valueAsNumber)
			? visioSizePositionCommand(selection.page, selection.id, field, input.valueAsNumber)
			: undefined;
		// Nothing to write (unchanged or refused): the box goes back to the shape's size.
		if (!command?.length) {
			this.render(state, true);
			return;
		}
		try {
			await this.controller.applySelectionEdits(command);
		} catch (error) {
			if (!isEditCancellation(error)) this.render(this.controller.state, true);
		}
		this.render(this.controller.state, true);
	}
	render(state: ViewerState, committed = false): void {
		const selection = this.#selection(state);
		const values = selection ? visioSizePositionState(selection.page, selection.id) : undefined;
		const editable = !!values && canEditDiagram(state) && !state.loading && !state.edit.busy;
		for (const input of this.inputs()) {
			const field = input.dataset.inlineSize as InlineField;
			// A box being typed in keeps what the user typed until it is committed.
			if (committed || this.root.activeElement !== input)
				input.value = values ? String(values[field]) : '';
			input.disabled = !editable;
		}
	}
}
