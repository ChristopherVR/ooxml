import type { VisioGeometryEdit } from 'ooxml-core/visio';
import type { ViewerController, ViewerState } from './controller.js';

export const geometryControlsTemplate = `<fieldset data-geometry><legend>Geometry (experimental)</legend><p>Pin X/Y use drawing inches, bottom-left origin, Y up. Resize holds the rotation pin fixed. Core rejects unsupported formulas, protected cells, groups, masters and referenced deletion. Native Visio reopen fidelity is unverified.</p><label>New rectangle ID <input data-geometry-field="id" type="text" maxlength="256"></label><label>Pin X (inches) <input data-geometry-field="x" type="number" step="any"></label><label>Pin Y (inches) <input data-geometry-field="y" type="number" step="any"></label><label>Width (inches) <input data-geometry-field="width" type="number" min="0" step="any"></label><label>Height (inches) <input data-geometry-field="height" type="number" min="0" step="any"></label><div class="edit-actions"><button type="button" data-geometry-action="create-rectangle">Create rectangle</button><button type="button" data-geometry-action="move-shape">Move selected</button><button type="button" data-geometry-action="resize-shape">Resize selected</button><button type="button" data-geometry-action="delete-shape">Delete selected</button></div><p data-geometry-error role="alert" hidden></p></fieldset>`;

/** Collect explicit coordinates; never infer a Visio pin from rendered SVG transforms. */
export class ViewerGeometryControls {
	#fields: Record<'id' | 'x' | 'y' | 'width' | 'height', HTMLInputElement>;
	#buttons: HTMLButtonElement[];
	#error: HTMLElement;
	#identity = '';
	#request = 0;
	constructor(
		private readonly panel: HTMLElement,
		private readonly controller: ViewerController,
	) {
		const field = (name: string) =>
			panel.querySelector<HTMLInputElement>(`[data-geometry-field="${name}"]`)!;
		this.#fields = {
			id: field('id'),
			x: field('x'),
			y: field('y'),
			width: field('width'),
			height: field('height'),
		};
		this.#buttons = [...panel.querySelectorAll<HTMLButtonElement>('[data-geometry-action]')];
		this.#error = panel.querySelector('[data-geometry-error]')!;
	}
	wire(): () => void {
		const Abort = this.panel.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		for (const field of Object.values(this.#fields))
			field.addEventListener('input', () => this.render(this.controller.state), {
				signal: events.signal,
			});
		for (const button of this.#buttons)
			button.addEventListener(
				'click',
				() => {
					void this.#run(button);
				},
				{ signal: events.signal },
			);
		return () => {
			++this.#request;
			events.abort();
		};
	}
	async #run(button: HTMLButtonElement): Promise<void> {
		if (button.disabled) return;
		const state = this.controller.state;
		const pageId = state.selectedShape?.pageId ?? state.document?.pages[state.pageIndex]?.id;
		if (pageId === undefined) return;
		const type = button.dataset.geometryAction as VisioGeometryEdit['type'];
		const shapeId = type === 'create-rectangle' ? this.#fields.id.value : state.selectedShape?.id;
		if (!shapeId) return;
		const x = this.#fields.x.valueAsNumber,
			y = this.#fields.y.valueAsNumber;
		const width = this.#fields.width.valueAsNumber,
			height = this.#fields.height.valueAsNumber;
		const target = { pageId, shapeId };
		const command: VisioGeometryEdit =
			type === 'create-rectangle'
				? { type, ...target, x, y, width, height }
				: type === 'move-shape'
					? { type, ...target, x, y }
					: type === 'resize-shape'
						? { type, ...target, width, height }
						: { type: 'delete-shape', ...target };
		const request = ++this.#request;
		this.#error.hidden = true;
		try {
			await this.controller.applyEdits([command]);
		} catch (error) {
			if (request === this.#request) {
				this.#error.textContent = error instanceof Error ? error.message : String(error);
				this.#error.hidden = false;
			}
		}
	}
	reset(): void {
		++this.#request;
		this.#error.hidden = true;
		for (const field of Object.values(this.#fields)) field.value = '';
	}
	render(state: ViewerState): void {
		const pageId = state.selectedShape?.pageId ?? state.document?.pages[state.pageIndex]?.id;
		const identity = JSON.stringify([
			this.controller.documentGeneration,
			state.pageIndex,
			pageId,
			state.selectedShape?.id,
		]);
		if (identity !== this.#identity) {
			this.#identity = identity;
			++this.#request;
			this.#error.hidden = true;
			for (const field of Object.values(this.#fields)) field.value = '';
		}
		const available =
			state.edit.sourceAvailable && pageId !== undefined && !state.loading && !state.edit.busy;
		for (const field of Object.values(this.#fields)) field.disabled = !available;
		const position = [this.#fields.x, this.#fields.y].every((field) =>
			Number.isFinite(field.valueAsNumber),
		);
		const size = [this.#fields.width, this.#fields.height].every(
			(field) => Number.isFinite(field.valueAsNumber) && field.valueAsNumber > 0,
		);
		for (const button of this.#buttons) {
			const action = button.dataset.geometryAction;
			button.disabled =
				!available ||
				(action === 'create-rectangle'
					? !this.#fields.id.value || !position || !size
					: !state.selectedShape ||
						(action === 'move-shape' && !position) ||
						(action === 'resize-shape' && !size));
		}
	}
}
