import type { ViewerController, ViewerState } from './controller';
import { editErrorMessage } from 'ooxml-core/visio/ui';

type Dialog = HTMLElement & { open: boolean; show(): void; close(): void };
type Button = HTMLElement & { disabled: boolean };
/** Rename draft uses the shared modal; page/name/reference semantics stay in core. */
export class ViewerPageRename {
	readonly dialog: Dialog;
	readonly input: HTMLInputElement;
	readonly error: HTMLElement;
	#buttons: Button[] = [];
	#generation = -1;
	#pageId = '';
	constructor(
		root: ShadowRoot,
		private readonly controller: ViewerController,
	) {
		const doc = root.ownerDocument;
		this.dialog = doc.createElement('office-ui-dialog') as Dialog;
		this.dialog.className = 'page-rename-dialog';
		this.dialog.setAttribute('heading', 'Rename Page');
		const label = doc.createElement('label');
		label.textContent = 'Page name';
		this.input = doc.createElement('input');
		this.input.type = 'text';
		this.input.maxLength = 255;
		this.input.setAttribute('aria-label', 'Page name');
		this.input.addEventListener('keydown', (event) => {
			if (event.key === 'Enter') {
				event.preventDefault();
				void this.#apply();
			}
		});
		label.append(this.input);
		this.error = doc.createElement('p');
		this.error.setAttribute('role', 'alert');
		this.dialog.append(label, this.error);
		for (const name of ['OK', 'Cancel']) {
			const button = doc.createElement('office-ui-button') as Button;
			button.slot = 'footer';
			button.setAttribute('label', name);
			button.setAttribute('command', `page-rename-${name.toLowerCase()}`);
			button.addEventListener('office-command', () => {
				if (name === 'Cancel') this.close();
				else void this.#apply();
			});
			this.#buttons.push(button);
			this.dialog.append(button);
		}
		root.append(this.dialog);
	}
	show(): void {
		const state = this.controller.state,
			page = state.document?.pages[state.pageIndex];
		if (!page || !state.edit.sourceAvailable || state.edit.busy || state.loading) return;
		this.#pageId = page.id;
		this.#generation = this.controller.documentGeneration;
		this.input.value = page.name;
		this.error.textContent = '';
		this.dialog.show();
		this.render(state);
		this.input.focus();
		this.input.select();
	}
	close(): void {
		this.dialog.close();
	}
	render(state: ViewerState): void {
		if (!this.dialog.open) return;
		if (
			!state.edit.sourceAvailable ||
			state.loading ||
			this.controller.documentGeneration !== this.#generation
		) {
			this.close();
			return;
		}
		this.input.disabled = state.edit.busy;
		for (const button of this.#buttons) button.disabled = state.edit.busy;
	}
	async #apply(): Promise<void> {
		if (
			!this.dialog.open ||
			this.input.disabled ||
			this.controller.documentGeneration !== this.#generation
		)
			return;
		const generation = this.#generation;
		try {
			await this.controller.applyEdits([
				{ type: 'rename-page', pageId: this.#pageId, name: this.input.value },
			]);
			if (this.#generation === generation) this.close();
		} catch (error) {
			if (
				this.dialog.open &&
				this.#generation === generation &&
				this.controller.documentGeneration === generation
			)
				this.error.textContent = editErrorMessage(error);
		}
	}
}
