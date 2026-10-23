import type { SmartArtLayout } from 'ooxml-core/pptx';

import { defineSmartArtPicker } from '../../smartart/smartart-picker';
import type { OfficeSmartArtPreview, OfficeUiSmartArtPicker } from '../../smartart/smartart-picker';
import { DIAGRAM_INSERT_CATEGORIES, DIAGRAM_INSERT_PRESETS } from 'ooxml-core/diagram';
import {
	smartArtCategoryDescriptionKey,
	smartArtCategoryLabelKey,
	smartArtPresetLabelKey,
} from '../render/smart-art-presets';
import { createInsertDialogFrame, insertDialogText } from './insert-dialog-frame';
import type { InsertDialogFrame, InsertDialogTranslate } from './insert-dialog-frame';

/** What Insert hands the host: the layout and the node texts to seed it with. */
export interface InsertSmartArtDialogDetail {
	layout: SmartArtLayout;
	items: string[];
}
export type InsertSmartArtDialogInsertEvent = CustomEvent<InsertSmartArtDialogDetail>;

export interface PptxUiInsertSmartArtDialogElement extends HTMLElement {
	open: boolean;
	translateLabel: InsertDialogTranslate | undefined;
	/** Replaces the live previews (tests, or a host that draws its own). */
	drawPreview: OfficeSmartArtPreview | undefined;
}

type LivePreview = typeof import('../dom/smartart-insert-preview');

/**
 * `pptx-ui-insert-smartart-dialog`: PowerPoint's Insert SmartArt dialog, the shared
 * `office-ui-smartart-picker` (categories, layout gallery, description pane and the node texts)
 * in the shared dialog. Tiles start as layout-family drawings and switch to live renders of the
 * diagram each preset inserts once the slide renderer has loaded (on first open). Controlled: the
 * host sets `open` (and `translateLabel`); Insert, a double click or Enter on a tile emits
 * `insert-smartart` with {@link InsertSmartArtDialogDetail}, and every close emits
 * `dialog-close` so the host clears its open state. The host inserts the diagram.
 */
export function definePptxInsertSmartArtDialog(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-insert-smartart-dialog')) {
		return;
	}
	defineSmartArtPicker(registry);

	class InsertSmartArtDialog extends HTMLElement implements PptxUiInsertSmartArtDialogElement {
		static observedAttributes = ['open'];
		#frame: InsertDialogFrame | undefined;
		#picker: OfficeUiSmartArtPicker | undefined;
		#translate: InsertDialogTranslate | undefined;
		#draw: OfficeSmartArtPreview | undefined;
		#live: Promise<void> | undefined;

		get open(): boolean {
			return this.hasAttribute('open');
		}
		set open(value: boolean) {
			this.toggleAttribute('open', Boolean(value));
		}
		get translateLabel(): InsertDialogTranslate | undefined {
			return this.#translate;
		}
		set translateLabel(value: InsertDialogTranslate | undefined) {
			if (value === this.#translate) {
				return;
			}
			this.#translate = value;
			this.#sync();
		}
		get drawPreview(): OfficeSmartArtPreview | undefined {
			return this.#draw;
		}
		set drawPreview(value: OfficeSmartArtPreview | undefined) {
			this.#draw = value;
			if (this.#picker) {
				this.#picker.drawPreview = value;
			}
		}

		connectedCallback(): void {
			this.#build();
			this.#sync();
			this.#show();
		}

		attributeChangedCallback(): void {
			if (this.isConnected) {
				this.#show();
			}
		}

		#text(key: string, fallback: string): string {
			return insertDialogText(this.#translate, key, fallback);
		}

		#build(): void {
			if (this.#frame) {
				return;
			}
			const picker = this.ownerDocument.createElement(
				'office-ui-smartart-picker',
			) as OfficeUiSmartArtPicker;
			picker.itemsEditor = true;
			picker.drawPreview = this.#draw;
			picker.addEventListener('office-smartart-activate', () => this.#insert());
			picker.addEventListener('office-smartart-change', () => this.#syncFooter());
			this.#picker = picker;
			this.#frame = createInsertDialogFrame(
				this,
				picker,
				() => this.#insert(),
				() => this.#closed(),
			);
		}

		#sync(): void {
			const picker = this.#picker;
			if (!picker || !this.#frame) {
				return;
			}
			picker.categories = DIAGRAM_INSERT_CATEGORIES.map((category) => ({
				id: category.id,
				label: this.#text(smartArtCategoryLabelKey(category.id), category.label),
				description: this.#text(smartArtCategoryDescriptionKey(category.id), category.description),
			}));
			picker.layouts = DIAGRAM_INSERT_PRESETS.map((preset) => ({
				id: preset.layout,
				category: preset.category,
				label: this.#text(smartArtPresetLabelKey(preset.layout), preset.label),
				family: preset.family,
				defaultItems: preset.defaultItems,
			}));
			picker.labels = {
				categories: this.#text('pptx.insertSmartArt.categories', 'SmartArt categories'),
				layouts: this.#text('pptx.insertSmartArt.layouts', 'SmartArt layouts'),
				items: this.#text('pptx.insertSmartArt.itemsLabel', 'Items (one per line)'),
			};
			this.#syncFooter();
		}

		#syncFooter(): void {
			if (!this.#picker || !this.#frame) {
				return;
			}
			this.#frame.sync(
				this.#text('pptx.smartart.insertTitle', 'Insert SmartArt'),
				this.#text('pptx.smartart.close', 'Close'),
				this.#text('pptx.smartart.cancel', 'Cancel'),
				this.#text('pptx.smartart.insert', 'Insert'),
				this.#picker.value.layout !== '',
			);
		}

		/** Loads the slide renderer once and switches the tiles to live previews. */
		#loadLivePreviews(): void {
			if (this.#live || this.#draw) {
				return;
			}
			this.#live = import('../dom/smartart-insert-preview')
				.then((module: LivePreview) => {
					if (this.#draw || !this.#picker) {
						return;
					}
					const t = (key: string) => this.#text(key, key);
					this.#picker.drawPreview = (layout, size) =>
						module.buildSmartArtInsertPreview(
							this.ownerDocument,
							t,
							layout.id as SmartArtLayout,
							size,
						);
				})
				.catch(() => undefined);
		}

		#show(): void {
			const frame = this.#frame;
			const picker = this.#picker;
			if (!frame || !picker || frame.dialog.open === this.open) {
				return;
			}
			if (this.open) {
				const first = DIAGRAM_INSERT_PRESETS[0];
				picker.value = { layout: first?.layout ?? '', items: [...(first?.defaultItems ?? [])] };
				this.#loadLivePreviews();
			}
			frame.dialog.open = this.open;
			if (this.open) {
				picker.shadowRoot
					?.querySelector<HTMLElement>('[role="option"][aria-selected="true"]')
					?.focus();
			}
		}

		#insert(): void {
			const value = this.#picker?.value;
			const preset = DIAGRAM_INSERT_PRESETS.find((entry) => entry.layout === value?.layout);
			if (!value || !preset) {
				return;
			}
			const items = value.items.length > 0 ? value.items : [...preset.defaultItems];
			this.dispatchEvent(
				new CustomEvent<InsertSmartArtDialogDetail>('insert-smartart', {
					detail: { layout: preset.layout, items },
					bubbles: true,
					composed: true,
				}),
			);
			if (this.#frame) {
				this.#frame.dialog.open = false;
			}
			this.#closed();
		}

		#closed(): void {
			this.removeAttribute('open');
			this.dispatchEvent(new CustomEvent('dialog-close', { bubbles: true, composed: true }));
		}
	}

	registry.define('pptx-ui-insert-smartart-dialog', InsertSmartArtDialog);
}
