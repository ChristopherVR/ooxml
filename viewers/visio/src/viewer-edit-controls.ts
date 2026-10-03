import { geometryControlsTemplate, ViewerGeometryControls } from './viewer-geometry-controls.js';
import type { ViewerController, ViewerState } from './controller.js';
import { selectedShape } from './shape-inspector.js';

/** Static markup only. Document text is assigned exclusively through value/textContent. */
export const editControlsTemplate = `<details class="edit-controls"><summary>Edit diagram (experimental)</summary><p id="edit-warning">Experimental source-backed editing. Unsupported targets are rejected. Native Visio compatibility is not verified. Keep your original file.</p><p data-edit-target></p><label for="edit-text">Selected shape text</label><textarea id="edit-text" rows="3" aria-describedby="edit-warning edit-target edit-status"></textarea><div class="edit-actions"><button type="button" data-edit="apply">Apply text</button><button type="button" data-edit="cancel">Cancel</button><button type="button" data-edit="undo">Undo</button><button type="button" data-edit="redo">Redo</button></div><p id="edit-status" role="status" aria-live="polite"></p><p data-edit-error role="alert" hidden></p><ul data-edit-diagnostics></ul>${geometryControlsTemplate}</details>`;

export class ViewerEditControls {
	readonly panel: HTMLDetailsElement;
	readonly input: HTMLTextAreaElement;
	#target: HTMLElement;
	#status: HTMLElement;
	#error: HTMLElement;
	#diagnostics: HTMLUListElement;
	#buttons: Record<'apply' | 'cancel' | 'undo' | 'redo', HTMLButtonElement>;
	#identity = '';
	#initial = '';
	#pageId: string | undefined;
	#shapeId: string | undefined;
	#request = 0;
	#pending = false;
	#localError: string | undefined;
	#controller: ViewerController;
	#geometry: ViewerGeometryControls;
	constructor(root: ShadowRoot, controller: ViewerController) {
		this.#controller = controller;
		this.panel = root.querySelector('.edit-controls')!;
		this.#geometry = new ViewerGeometryControls(this.panel, controller);
		this.input = this.panel.querySelector('textarea')!;
		this.#target = this.panel.querySelector('[data-edit-target]')!;
		this.#target.id = 'edit-target';
		this.#status = this.panel.querySelector('#edit-status')!;
		this.#error = this.panel.querySelector('[data-edit-error]')!;
		this.#diagnostics = this.panel.querySelector('[data-edit-diagnostics]')!;
		const button = (name: string) =>
			this.panel.querySelector<HTMLButtonElement>(`[data-edit="${name}"]`)!;
		this.#buttons = {
			apply: button('apply'),
			cancel: button('cancel'),
			undo: button('undo'),
			redo: button('redo'),
		};
	}
	wire(): () => void {
		const disposeGeometry = this.#geometry.wire();
		const Abort = this.panel.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.input.addEventListener('input', () => this.render(this.#controller.state), options);
		this.input.addEventListener(
			'keydown',
			(event) => {
				if (event.isComposing) return;
				if (event.key === 'Escape') {
					event.preventDefault();
					this.reset();
				}
				if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
					event.preventDefault();
					this.#buttons.apply.click();
				}
			},
			options,
		);
		this.#buttons.cancel.addEventListener('click', () => this.reset(), options);
		for (const action of ['apply', 'undo', 'redo'] as const)
			this.#buttons[action].addEventListener(
				'click',
				() => {
					void this.#run(action);
				},
				options,
			);
		return () => {
			disposeGeometry();
			events.abort();
			// Disposal must also work after a consumer has destroyed the controller.
			// The element owns transaction cancellation; only forget local draft state here.
			++this.#request;
			this.#pending = false;
			this.#localError = undefined;
			this.#identity = '';
			this.input.value = this.#initial;
		};
	}
	reset(focus = true): void {
		this.#geometry.reset();
		++this.#request;
		this.#pending = false;
		this.#localError = undefined;
		this.#controller.cancelEdit();
		this.input.value = this.#initial;
		this.render(this.#controller.state);
		if (focus && !this.input.disabled) this.input.focus();
	}
	async #run(action: 'apply' | 'undo' | 'redo'): Promise<void> {
		if (this.#buttons[action].disabled) return;
		const request = ++this.#request;
		this.#pending = true;
		this.#localError = undefined;
		this.render(this.#controller.state);
		try {
			if (action === 'apply') {
				if (this.#pageId === undefined || this.#shapeId === undefined) return;
				await this.#controller.replacePlainText(this.#pageId, this.#shapeId, this.input.value);
			} else await this.#controller[action]();
		} catch (error) {
			if (request === this.#request)
				this.#localError = error instanceof Error ? error.message : String(error);
		} finally {
			if (request === this.#request) {
				this.#pending = false;
				this.render(this.#controller.state);
			}
		}
	}
	render(state: ViewerState): void {
		this.#geometry.render(state);
		const selection = state.selectedShape;
		const pageId = selection?.pageId ?? state.document?.pages[state.pageIndex]?.id;
		const shape = selectedShape(state.document, selection, state.pageIndex);
		const identity = JSON.stringify([
			this.#controller.documentGeneration,
			state.pageIndex,
			pageId,
			selection?.id,
		]);
		if (identity !== this.#identity) {
			this.#identity = identity;
			++this.#request;
			this.#pending = false;
			this.#localError = undefined;
			this.#pageId = shape ? pageId : undefined;
			this.#shapeId = shape?.id;
			this.#initial = shape?.text?.plainText ?? '';
			this.input.value = this.#initial;
		}
		const busy = state.loading || state.edit.busy || this.#pending;
		const available = state.edit.sourceAvailable && !!shape && pageId !== undefined;
		this.panel.setAttribute('aria-busy', String(busy));
		this.#target.textContent = shape
			? `Page ID: ${pageId} · Shape ID: ${shape.id} · ${shape.name || 'Unnamed shape'}`
			: 'Select a shape to replace its plain text.';
		this.input.disabled = !available || busy;
		this.#buttons.apply.disabled = !available || busy || this.input.value === this.#initial;
		this.#buttons.cancel.disabled =
			!state.edit.busy && !this.#pending && this.input.value === this.#initial;
		this.#buttons.undo.disabled = busy || !state.edit.canUndo;
		this.#buttons.redo.disabled = busy || !state.edit.canRedo;
		this.#status.textContent = state.loading
			? 'Opening diagram…'
			: busy
				? 'Updating diagram…'
				: !state.edit.sourceAvailable
					? 'Open a .vsdx file to edit or download a VSDX copy. Model-only documents cannot be saved.'
					: `${state.edit.dirty ? 'Edited copy' : 'Original bytes'}${state.edit.historyTruncated ? ' · Earlier undo history was discarded.' : ''}`;
		const error = this.#localError ?? state.edit.error?.message;
		this.#error.hidden = !error;
		this.#error.textContent = error ?? '';
		this.#diagnostics.replaceChildren(
			...state.edit.diagnostics.map((diagnostic) => {
				const item = this.panel.ownerDocument.createElement('li');
				item.textContent = `${diagnostic.code}: ${diagnostic.message}`;
				return item;
			}),
		);
	}
}
