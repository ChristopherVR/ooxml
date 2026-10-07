import type { ViewerController, ViewerState } from './controller';
import { editErrorMessage } from 'ooxml-core/visio/ui';
import type { VisioPageReorder } from 'ooxml-core/visio';

type Dialog = HTMLElement & { open: boolean; show(): void; close(): void };
type Button = HTMLElement & { disabled: boolean };
/** Page-order draft on the shared modal. Core applies the complete permutation atomically. */
export class ViewerPageOrder {
	readonly dialog: Dialog;
	readonly list: HTMLSelectElement;
	readonly error: HTMLElement;
	#buttons: Button[] = [];
	#generation = -1;
	constructor(
		root: ShadowRoot,
		private readonly controller: ViewerController,
	) {
		const doc = root.ownerDocument;
		this.dialog = doc.createElement('office-ui-dialog') as Dialog;
		this.dialog.className = 'page-order-dialog';
		this.dialog.setAttribute('heading', 'Reorder Pages');
		const label = doc.createElement('label');
		label.textContent = 'Page order';
		this.list = doc.createElement('select');
		this.list.size = 8;
		this.list.setAttribute('aria-label', 'Page order');
		this.list.addEventListener('change', () => this.render(this.controller.state));
		label.append(this.list);
		this.error = doc.createElement('p');
		this.error.setAttribute('role', 'alert');
		const button = (label: string, action: () => void, footer = false) => {
			const item = doc.createElement('office-ui-button') as Button;
			item.setAttribute('label', label);
			item.setAttribute('command', `page-order-${label.toLowerCase().replaceAll(' ', '-')}`);
			if (footer) item.slot = 'footer';
			item.addEventListener('office-command', action);
			this.#buttons.push(item);
			this.dialog.append(item);
		};
		this.dialog.append(label);
		button('Move Up', () => this.#move(-1));
		button('Move Down', () => this.#move(1));
		this.dialog.append(this.error);
		button(
			'OK',
			() => {
				void this.#apply();
			},
			true,
		);
		button('Cancel', () => this.close(), true);
		root.append(this.dialog);
	}
	show(): void {
		const state = this.controller.state;
		if (!state.document || !state.edit.sourceAvailable || state.edit.busy || state.loading) return;
		this.list.replaceChildren(
			...state.document.pages.map((page) => {
				const option = this.list.ownerDocument.createElement('option');
				option.value = page.id;
				option.textContent = page.name;
				return option;
			}),
		);
		this.list.selectedIndex = state.pageIndex;
		this.error.textContent = '';
		this.#generation = this.controller.documentGeneration;
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
		this.list.disabled = state.edit.busy;
		for (const button of this.#buttons) button.disabled = state.edit.busy;
		this.#buttons[0]!.disabled ||= this.list.selectedIndex <= 0;
		this.#buttons[1]!.disabled ||= this.list.selectedIndex >= this.list.options.length - 1;
	}
	#move(direction: -1 | 1): void {
		if (this.list.disabled) return;
		const index = this.list.selectedIndex,
			target = index + direction;
		if (index < 0 || target < 0 || target >= this.list.options.length) return;
		const selected = this.list.options[index]!;
		this.list.insertBefore(
			selected,
			direction < 0 ? this.list.options[target]! : (this.list.options[target + 1] ?? null),
		);
		selected.selected = true;
		this.render(this.controller.state);
	}
	async #apply(): Promise<void> {
		if (this.list.disabled || this.controller.documentGeneration !== this.#generation) return;
		const generation = this.#generation;
		try {
			const commands: VisioPageReorder[] = Array.from(this.list.options, (option, index) => ({
				type: 'reorder-page',
				pageId: option.value,
				index,
			}));
			await this.controller.applyEdits(commands);
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
