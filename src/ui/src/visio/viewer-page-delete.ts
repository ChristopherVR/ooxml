import type { ViewerController, ViewerState } from './controller.js';
import { editErrorMessage } from 'ooxml-core/visio/ui';
type Dialog = HTMLElement & { open: boolean; show(): void; close(): void };
type Button = HTMLElement & { disabled: boolean };
/** Confirm deleting a stable page ID; core owns dependencies and replacement-page semantics. */
export class ViewerPageDelete {
	readonly dialog: Dialog;
	readonly message: HTMLElement;
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
		this.dialog.className = 'page-delete-dialog';
		this.dialog.setAttribute('heading', 'Delete Page');
		this.message = doc.createElement('p');
		this.error = doc.createElement('p');
		this.error.setAttribute('role', 'alert');
		this.dialog.append(this.message, this.error);
		for (const label of ['Delete', 'Cancel']) {
			const button = doc.createElement('office-ui-button') as Button;
			button.slot = 'footer';
			button.setAttribute('label', label);
			button.setAttribute('command', `page-delete-${label.toLowerCase()}`);
			button.addEventListener('office-command', () => {
				if (label === 'Cancel') this.close();
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
		this.message.textContent = `Delete “${page.name}” and its shapes? You can undo this change. A blank page remains when the last foreground page is deleted.`;
		this.error.textContent = '';
		this.dialog.show();
		this.render(state);
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
		for (const button of this.#buttons) button.disabled = state.edit.busy;
	}
	async #apply(): Promise<void> {
		if (
			!this.dialog.open ||
			this.controller.state.edit.busy ||
			this.controller.documentGeneration !== this.#generation
		)
			return;
		const generation = this.#generation;
		try {
			await this.controller.applyEdits([{ type: 'delete-page', pageId: this.#pageId }]);
			if (generation === this.#generation) this.close();
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
