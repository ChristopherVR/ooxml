import {
	editErrorMessage,
	isEditCancellation,
	visioDrawBounds,
	visioDrawPlan,
	type VisioDrawingPoint,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { ViewerDrawingGesture, type DrawingGesture } from './viewer-drawing-gesture';
import { createTextDraft } from './viewer-text-draft';

interface Draft {
	drag: DrawingGesture;
	end: VisioDrawingPoint;
	view: ReturnType<typeof createTextDraft>;
	pending: boolean;
}
/** Fixed-size plain-text boxes. Native automatic sizing and rich content editing are separate. */
export class ViewerTextTool {
	#gesture: ViewerDrawingGesture;
	#draft: Draft | undefined;
	#request = 0;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: {
			active(): boolean;
			announce(message: string): void;
			revealEdit(): void;
		},
	) {
		this.#gesture = new ViewerDrawingGesture(viewport, controller, {
			tool: () => (options.active() && !this.#draft ? 'text' : undefined),
			announce: options.announce,
			existingShape: (event) => this.#existingShape(event),
			finish: async (drag, end) => this.#begin(drag, end),
		});
	}
	get drafting(): boolean {
		return this.#gesture.drawing || !!this.#draft;
	}
	wire(): () => void {
		const dispose = this.#gesture.wire();
		return () => {
			this.cancel();
			dispose();
		};
	}
	render(state: ViewerState): void {
		this.#gesture.render(state);
		const draft = this.#draft;
		if (!draft) return;
		const replaced =
			state.loading ||
			!state.edit.sourceAvailable ||
			this.controller.sourceGeneration !== draft.drag.sourceGeneration;
		const pendingChanged =
			draft.pending &&
			state.edit.busy &&
			(state.document !== draft.drag.document ||
				state.document?.pages[state.pageIndex] !== draft.drag.page ||
				state.selectedShapes !== draft.drag.selection ||
				state.zoom !== draft.drag.zoom ||
				state.layerVisibilityOverrides !== draft.drag.layers);
		if (
			!this.options.active() ||
			replaced ||
			pendingChanged ||
			(!draft.pending && !this.#current(draft))
		)
			this.cancel();
	}
	cancel(): void {
		++this.#request;
		this.#gesture.cancel();
		const draft = this.#draft;
		this.#draft = undefined;
		draft?.view.dispose();
		draft?.drag.preview.remove();
		if (
			draft?.pending &&
			!this.controller.state.loading &&
			this.controller.state.document === draft.drag.document &&
			this.controller.state.edit.busy
		)
			this.controller.cancelEdit();
	}
	#current(draft: Draft): boolean {
		const state = this.controller.state;
		return (
			this.options.active() &&
			draft.drag.svg.isConnected &&
			state.document === draft.drag.document &&
			state.document?.pages[state.pageIndex] === draft.drag.page &&
			this.controller.isCreationTokenCurrent(draft.drag.token)
		);
	}
	#existingShape(event: PointerEvent): boolean {
		const group = (event.target as Element)?.closest?.<SVGGElement>('[data-shape-id]');
		if (!group) return false;
		const state = this.controller.state,
			page = state.document?.pages[state.pageIndex];
		if (!page || group.dataset.pageId !== page.id) return true;
		event.preventDefault();
		event.stopImmediatePropagation();
		this.controller.selectShape({
			id: group.dataset.shapeId!,
			name: group.dataset.shapeName ?? '',
			pageId: page.id,
		});
		if (
			this.controller.state.selectedShape?.id === group.dataset.shapeId &&
			this.controller.state.selectedShape?.pageId === page.id &&
			this.controller.state.document?.pages[this.controller.state.pageIndex] === page &&
			this.controller.state.document === state.document
		)
			this.options.revealEdit();
		return true;
	}
	#begin(drag: DrawingGesture, end: VisioDrawingPoint): void {
		if (!this.#gesture.current(drag)) return;
		const bounds = visioDrawBounds(drag.start, end);
		const view = createTextDraft(
			this.viewport,
			drag.svg,
			bounds,
			() => void this.#apply(),
			() => this.cancel(),
		);
		if (!this.#gesture.current(drag)) {
			view.dispose();
			return;
		}
		drag.preview.classList.add('text-box-draft-frame');
		for (const [name, value] of Object.entries({
			x: bounds.x,
			y: bounds.y,
			width: bounds.width,
			height: bounds.height,
		}))
			drag.preview.setAttribute(name, String(value));
		drag.svg.append(drag.preview);
		this.#draft = { drag, end, view, pending: false };
	}
	async #apply(): Promise<void> {
		const draft = this.#draft;
		if (!draft || draft.pending || !this.#current(draft)) return;
		const text = draft.view.input.value;
		if (text === '') {
			this.cancel();
			return;
		}
		const request = ++this.#request;
		try {
			const command = visioDrawPlan(draft.drag.page, 'text', draft.drag.start, draft.end, text);
			if (!command || !this.#current(draft)) return;
			draft.pending = true;
			draft.view.setBusy(true);
			await this.controller.applyCreationEdits([command], draft.drag.token);
			if (request !== this.#request) return;
			this.#draft = undefined;
			draft.view.dispose();
			draft.drag.preview.remove();
			this.options.announce(`Text box ${command.shapeId} added.`);
			this.viewport.focus({ preventScroll: true });
		} catch (error) {
			if (request !== this.#request) return;
			this.#draft = undefined;
			draft.view.dispose();
			draft.drag.preview.remove();
			if (!isEditCancellation(error)) this.options.announce(editErrorMessage(error));
		}
	}
}
