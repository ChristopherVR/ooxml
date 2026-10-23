import { html, nothing } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import { parseSvgPreview } from '../ribbon/safe-svg';
import css from './smartart-picker.css?raw';
import { smartArtLayoutTileSvg } from './tiles';

/** A category of the list on the left. `description` is shown under the chosen layout. */
export interface OfficeSmartArtCategory {
	id: string;
	label: string;
	description?: string | undefined;
}

/** One layout of the gallery. `family` picks the default tile drawing (list, process, ...). */
export interface OfficeSmartArtLayout {
	id: string;
	category: string;
	label: string;
	family?: string | undefined;
	/** The node texts the items editor starts from when the layout is chosen. */
	defaultItems?: readonly string[] | undefined;
}

/** What the picker edits: the chosen layout and, with the items editor, the node texts. */
export interface OfficeSmartArtValue {
	layout: string;
	items: string[];
}

/**
 * Draws a layout's preview: SVG markup (parsed with executable parts removed) or an element the
 * host built (a live render), or null for the default family tile. `tile` is true for a gallery
 * tile, false for the large preview beside it; `size` is the box in pixels.
 */
export type OfficeSmartArtPreview = (
	layout: OfficeSmartArtLayout,
	size: { width: number; height: number },
	tile: boolean,
) => string | Element | null | undefined;

export interface OfficeSmartArtPickerLabels {
	categories: string;
	layouts: string;
	items: string;
}

const ENGLISH: OfficeSmartArtPickerLabels = {
	categories: 'SmartArt categories',
	layouts: 'SmartArt layouts',
	items: 'Items (one per line)',
};

const TILE = { width: 64, height: 40 };
const PREVIEW = { width: 200, height: 125 };
const COLUMNS = 3;

const sameItems = (a: readonly string[], b: readonly string[]) =>
	a.length === b.length && a.every((item, index) => item === b[index]);

/**
 * `<office-ui-smartart-picker>`: the body of every product's Insert SmartArt dialog: a category
 * list, a gallery of layout tiles (`role="listbox"`), and a pane with the chosen layout's large
 * preview, name and its category's description. The product owns the dialog around it, the OK and
 * Cancel buttons and the diagram model. The host sets `categories`, `layouts` and `labels`
 * (translated), and optionally `drawPreview` (a live render; without it each tile draws its
 * layout family in `accent`). `items-editor` adds a text box of node texts, one per line, seeded
 * from the chosen layout's `defaultItems`. Read and write the choice with `value` and the shown
 * category with `category`; choosing a category selects its first layout.
 *
 * Emits `office-smartart-change` with the value on every choice or edit, and
 * `office-smartart-activate` with the value when a tile is double-clicked or Enter is pressed on
 * it (the host inserts, as OK would).
 */
export class OfficeUiSmartArtPicker extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		itemsEditor: { type: Boolean, attribute: 'items-editor' },
		accent: { type: String },
	};
	/** Shows the node texts editor. */
	declare itemsEditor: boolean;
	/** The fill of the default tiles (a CSS hex or rgb colour). */
	declare accent: string | null;

	private categoryList: readonly OfficeSmartArtCategory[] = [];
	private layoutList: readonly OfficeSmartArtLayout[] = [];
	private text: OfficeSmartArtPickerLabels = ENGLISH;
	private draw: OfficeSmartArtPreview | undefined;
	private shown = '';
	private choice: OfficeSmartArtValue = { layout: '', items: [] };
	private readonly previews = new Map<string, Element | null>();

	constructor() {
		super();
		this.itemsEditor = false;
		this.accent = null;
	}

	get categories(): readonly OfficeSmartArtCategory[] {
		return this.categoryList;
	}
	set categories(value: readonly OfficeSmartArtCategory[]) {
		this.categoryList = value;
		if (!value.some((category) => category.id === this.shown)) this.showCategory(value[0]?.id);
		this.requestUpdate('categories');
	}
	get layouts(): readonly OfficeSmartArtLayout[] {
		return this.layoutList;
	}
	set layouts(value: readonly OfficeSmartArtLayout[]) {
		this.layoutList = value;
		this.previews.clear();
		if (!value.some((layout) => layout.id === this.choice.layout)) this.showCategory(this.shown);
		this.requestUpdate('layouts');
	}
	get labels(): OfficeSmartArtPickerLabels {
		return this.text;
	}
	set labels(value: Partial<OfficeSmartArtPickerLabels>) {
		this.text = { ...ENGLISH, ...value };
		this.requestUpdate('labels');
	}
	get drawPreview(): OfficeSmartArtPreview | undefined {
		return this.draw;
	}
	set drawPreview(value: OfficeSmartArtPreview | undefined) {
		this.draw = value;
		this.previews.clear();
		this.requestUpdate('drawPreview');
	}
	/** The category whose layouts the gallery shows. */
	get category(): string {
		return this.shown;
	}
	set category(value: string) {
		this.showCategory(value);
		this.requestUpdate('category');
	}
	/** The choice; `items` leaves out blank lines and surrounding spaces. */
	get value(): OfficeSmartArtValue {
		const items = this.choice.items.map((item) => item.trim()).filter(Boolean);
		return { layout: this.choice.layout, items };
	}
	set value(value: OfficeSmartArtValue) {
		this.choice = { layout: value.layout, items: [...value.items] };
		const layout = this.layoutList.find((entry) => entry.id === value.layout);
		if (layout) this.shown = layout.category;
		this.requestUpdate('value');
	}

	/** Shows `id`'s layouts and selects the first, unless the choice is already one of them. */
	private showCategory(id: string | undefined): void {
		this.shown = id ?? '';
		const current = this.layoutList.find((layout) => layout.id === this.choice.layout);
		if (current?.category === this.shown) return;
		const first = this.layoutList.find((layout) => layout.category === this.shown);
		this.choice = { layout: first?.id ?? '', items: [...(first?.defaultItems ?? [])] };
	}

	private pick(layout: OfficeSmartArtLayout): void {
		const previous = this.layoutList.find((entry) => entry.id === this.choice.layout);
		// Text the user typed survives a change of layout; untouched defaults follow the layout.
		const edited = previous && !sameItems(this.choice.items, previous.defaultItems ?? []);
		this.choice = {
			layout: layout.id,
			items: edited ? this.choice.items : [...(layout.defaultItems ?? [])],
		};
		this.requestUpdate('value');
		this.fire('office-smartart-change', this.value);
	}

	private chooseCategory(id: string): void {
		const before = this.choice.layout;
		this.showCategory(id);
		this.requestUpdate('category');
		if (this.choice.layout !== before) this.fire('office-smartart-change', this.value);
	}

	private visible(): OfficeSmartArtLayout[] {
		return this.layoutList.filter((layout) => layout.category === this.shown);
	}

	private onTileKey(event: KeyboardEvent, index: number): void {
		const layouts = this.visible();
		const step: Record<string, number> = {
			ArrowRight: 1,
			ArrowLeft: -1,
			ArrowDown: COLUMNS,
			ArrowUp: -COLUMNS,
		};
		let next: number | undefined;
		if (event.key in step)
			next = Math.min(layouts.length - 1, Math.max(0, index + step[event.key]!));
		else if (event.key === 'Home') next = 0;
		else if (event.key === 'End') next = layouts.length - 1;
		else if (event.key === 'Enter') {
			event.preventDefault();
			this.fire('office-smartart-activate', this.value);
			return;
		}
		const layout = next === undefined ? undefined : layouts[next];
		if (!layout) return;
		event.preventDefault();
		this.pick(layout);
		[...this.renderRoot.querySelectorAll<HTMLElement>('[data-layout]')]
			.find((tile) => tile.dataset.layout === layout.id)
			?.focus();
	}

	private preview(layout: OfficeSmartArtLayout, tile: boolean): Element | null {
		const key = `${tile ? 't' : 'p'}:${layout.id}`;
		if (this.previews.has(key)) return this.previews.get(key) ?? null;
		const size = tile ? TILE : PREVIEW;
		let drawn: string | Element | null | undefined;
		try {
			drawn = this.draw?.(layout, size, tile);
		} catch {
			drawn = undefined;
		}
		const svg =
			drawn ??
			smartArtLayoutTileSvg(layout.family ?? layout.category, this.accent ?? undefined, size);
		const node = typeof svg === 'string' ? parseSvgPreview(this.ownerDocument, svg) : svg;
		this.previews.set(key, node);
		return node;
	}

	protected override render() {
		const text = this.text;
		const layouts = this.visible();
		const chosen = this.layoutList.find((layout) => layout.id === this.choice.layout);
		const category = this.categoryList.find((entry) => entry.id === this.shown);
		const focusable = layouts.some((layout) => layout.id === this.choice.layout)
			? this.choice.layout
			: layouts[0]?.id;
		return html`<nav class="categories" aria-label=${text.categories}
				>${this.categoryList.map(
					(entry) =>
						html`<button
							type="button"
							class="category"
							data-category=${entry.id}
							aria-pressed=${entry.id === this.shown ? 'true' : 'false'}
							@click=${() => this.chooseCategory(entry.id)}
							>${entry.label}</button
						>`,
				)}</nav
			><div class="layouts" role="listbox" aria-label=${text.layouts}
				>${layouts.map(
					(layout, index) =>
						html`<button
							type="button"
							class="layout"
							role="option"
							data-layout=${layout.id}
							title=${layout.label}
							aria-selected=${layout.id === this.choice.layout ? 'true' : 'false'}
							tabindex=${layout.id === focusable ? '0' : '-1'}
							@click=${() => this.pick(layout)}
							@dblclick=${() => {
								this.pick(layout);
								this.fire('office-smartart-activate', this.value);
							}}
							@keydown=${(event: KeyboardEvent) => this.onTileKey(event, index)}
							><span class="tile" aria-hidden="true">${this.preview(layout, true) ?? nothing}</span
							><span class="name">${layout.label}</span></button
						>`,
				)}</div
			><div class="details" aria-live="polite"
				>${
					chosen
						? html`<div class="preview" aria-hidden="true"
									>${this.preview(chosen, false) ?? nothing}</div
								><h3 class="title">${chosen.label}</h3
								>${category?.description ? html`<p class="description">${category.description}</p>` : nothing}`
						: nothing
				}${
					this.itemsEditor
						? html`<label class="items"
								><span>${text.items}</span
								><textarea
									data-key="items"
									rows="5"
									aria-label=${text.items}
									.value=${this.choice.items.join('\n')}
									@input=${(event: Event) => {
										const value = (event.target as HTMLTextAreaElement).value;
										this.choice = { ...this.choice, items: value.split('\n') };
										this.fire('office-smartart-change', this.value);
									}}
								></textarea>
							</label>`
						: nothing
				}</div
			>`;
	}
}

export const defineSmartArtPicker = definer(
	'office-ui-smartart-picker',
	() => OfficeUiSmartArtPicker,
);
