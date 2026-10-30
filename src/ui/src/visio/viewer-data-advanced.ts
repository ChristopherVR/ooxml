import type { VisioPage } from 'ooxml-core/visio';
import {
	visioDataGraphicFields,
	visioDataGraphicItems,
	visioDeleteDataGraphicItemEdits,
	type VisioDataGraphicItem,
} from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController } from './controller';
import { InsertDialog } from './viewer-insert-dialog';
import { itemLabel } from './viewer-data-labels';

type Run = (action: () => Promise<void>, success: string) => void;

/**
 * Advanced Data Graphics: the Edit Data Graphic dialog. It lists the items of the first selected
 * shape (a display type, a Shape Data field and a place each); New and Edit open the item form,
 * Delete removes the item from every selected shape that has it.
 */
export class ViewerAdvancedGraphics {
	readonly dialog: InsertDialog;
	#items: VisioDataGraphicItem[] = [];
	/** Opens the item form, empty for a new item. */
	onEdit: (item: VisioDataGraphicItem | undefined) => void = () => {};
	constructor(
		root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
		private readonly run: Run,
		private readonly targets: (page: VisioPage) => string[],
	) {
		this.dialog = new InsertDialog(
			root,
			'data-graphic-items-dialog',
			'Edit Data Graphic',
			[{ name: 'items', label: 'Items in this data graphic', choices: true }],
			['New Item', 'Edit Item', 'Delete Item', 'Close'],
			(button) => this.#press(button),
		);
		const list = this.dialog.fields.get('items') as HTMLSelectElement;
		list.size = 6;
		list.addEventListener('change', () => this.#sync());
		list.addEventListener('dblclick', () => this.#press('Edit Item'));
	}
	#page() {
		const state = this.controller.state;
		return state.document?.pages[state.pageIndex];
	}
	/** The shape whose items are listed: the first selected top-level shape. */
	#owner(page: VisioPage): string | undefined {
		return this.targets(page).find((id) => page.shapes.some((shape) => shape.id === id));
	}
	open(): void {
		const page = this.#page();
		if (!page || !canEditDiagram(this.controller.state)) return;
		const selected = this.controller.state.selectedShapes.some((shape) =>
			page.shapes.some((item) => item.id === shape.id),
		);
		if (!selected) return this.announce('Select a shape to edit its data graphic.');
		if (!visioDataGraphicFields(page).length)
			return this.announce(
				'Link or define Shape Data first; data graphics display Shape Data fields.',
			);
		this.refresh();
		this.dialog.error.textContent = '';
		this.dialog.dialog.show();
		this.dialog.fields.get('items')!.focus();
	}
	/** Rebuild the list from the drawing, keeping the selected item when it is still there. */
	refresh(): void {
		const page = this.#page();
		const owner = page && this.#owner(page);
		this.#items = page && owner ? visioDataGraphicItems(page, owner) : [];
		const fields = page ? visioDataGraphicFields(page) : [];
		const labelOf = (name: string) => fields.find((field) => field.name === name)?.label ?? name;
		const select = this.dialog.fields.get('items') as HTMLSelectElement;
		const previous = select.value;
		this.dialog.choices(
			'items',
			this.#items.map((_, index) => String(index)),
			Number(previous) < this.#items.length ? previous || '0' : '0',
			{
				labels: this.#items.map((item) => itemLabel(item, labelOf(item.field))),
				none: false,
				noneLabel: 'No items yet. Choose New Item.',
			},
		);
		if (!this.#items.length) select.replaceChildren();
		this.#sync();
	}
	#selected(): VisioDataGraphicItem | undefined {
		return this.#items[Number((this.dialog.fields.get('items') as HTMLSelectElement).value)];
	}
	#sync(): void {
		const has = !!this.#selected();
		for (const name of ['Edit Item', 'Delete Item']) this.dialog.buttons.get(name)!.disabled = !has;
	}
	#press(button: string): void {
		if (button === 'Close') return this.dialog.dialog.close();
		const page = this.#page();
		if (!page) return;
		if (button === 'New Item') return this.onEdit(undefined);
		const item = this.#selected();
		if (!item) return;
		if (button === 'Edit Item') return this.onEdit(item);
		const edits = visioDeleteDataGraphicItemEdits(page, this.targets(page), item);
		this.run(async () => {
			await this.controller.applyEdits(edits);
			this.refresh();
		}, 'Deleted the item from the data graphic.');
	}
}
