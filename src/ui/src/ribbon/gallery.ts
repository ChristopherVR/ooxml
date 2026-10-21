import { html, render, type TemplateResult } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { repeat } from 'lit/directives/repeat.js';
import { html as staticHtml, unsafeStatic } from 'lit/static-html.js';
import { OfficeElement, controlStyles } from '../base';
import { createIconSvg, paintIcon } from '../icons';
import { definer, present } from '../registry';
import { attachGalleryStyles } from './gallery-styles';
import { parseSvgPreview } from './safe-svg';

/** `action` is set when a tile's context menu entry was chosen instead of the tile itself. */
export type OfficeGalleryPickEvent = CustomEvent<{
	gallery: string;
	itemId: string;
	action?: string;
}>;
/** The tile under the pointer (or keyboard focus) of a gallery whose state sets `preview`. */
export type OfficeGalleryPreviewEvent = CustomEvent<{ gallery: string; itemId: string }>;
/** The preview ended: the pointer left the tiles, the popup closed or a pick was made. */
export type OfficeGalleryPreviewEndEvent = CustomEvent<{ gallery: string }>;

/** An entry of a tile's context menu (right-click, the Menu key or Shift+F10). */
export interface OfficeGalleryAction {
	id: string;
	label: string;
}

/** One tile. `preview` is SVG markup, parsed as an SVG document with executable parts removed. */
export interface OfficeGalleryItem {
	id: string;
	label: string;
	applied?: boolean | undefined;
	preview?: string | undefined;
	/** Context menu entries; a plain click still picks the tile itself. */
	actions?: readonly OfficeGalleryAction[] | undefined;
}

export interface OfficeGallerySection {
	title?: string | undefined;
	columns: number;
	tileWidth: number;
	tileHeight: number;
	items: readonly OfficeGalleryItem[];
}

/** A gallery, translated. With `command` it is a single button that picks the first item. */
export interface OfficeGalleryState {
	id: string;
	label: string;
	/** Trigger name in the inline strip; defaults to "More <label>". */
	moreLabel?: string | undefined;
	disabled?: boolean | undefined;
	command?: { icon: string; large?: boolean | undefined; hint?: string | undefined } | undefined;
	sections: readonly OfficeGallerySection[];
	/** Tiles shown inline (`mode="inline"`); defaults to the first section's items. */
	inline?: readonly OfficeGalleryItem[] | undefined;
	/**
	 * Emit `previewEvent` while a tile is hovered or keyboard-focused and `previewEndEvent` when
	 * that stops (Themes, Colors, ...: the host shows the result without committing it).
	 */
	preview?: boolean | undefined;
	/** Commands listed under the tiles of the popup ("Format Background..."); a click picks `id`. */
	footer?: readonly OfficeGalleryAction[] | undefined;
}

type Statics = {
	pickEvent: string;
	previewEvent: string;
	previewEndEvent: string;
	triggerAttribute: string;
	popupAttribute: string;
	itemAttribute: string;
	compactAttribute: string | null;
};

const HOST_CSS = ':host { display: inline-flex; position: relative; flex: none; }';

/**
 * `<office-ui-gallery>`: Office's ribbon gallery (Shape Styles, Themes, Table Styles). Set
 * `state` (`OfficeGalleryState`, translated). `mode="inline"` shows a strip of tiles beside a
 * More button; otherwise the trigger is a labelled dropdown (`chevron-only` keeps just the
 * chevron, `icon` names its glyph). The popup opens below the trigger, keeps inside the window
 * and closes on Escape, an outside press or a pick. `mode="panel"` shows labelled sections directly
 * for modal pickers, with arrow/Home/End selection and no dropdown trigger; `compact` shows its
 * tiles without visible labels, sharing the panel's width around previews at their size, and lets
 * arrow keys only move focus, for a pane whose pick applies at once. ArrowDown on the trigger
 * opens it and focuses the first tile. Dropdown arrows/Home/End move focus across sections
 * without applying a choice.
 * A pick emits `office-gallery-pick` `{ gallery, itemId }`.
 *
 * Tiles render in the light DOM beside the host's children (Lit renders the template into the
 * host itself) so products can query them (`data-gallery-item`); the rules join the host's root
 * once. Static `pickEvent`, `triggerAttribute`, `popupAttribute`, `itemAttribute` and
 * `compactAttribute` keep a product's published names; override `triggerIcon()` to draw a product
 * glyph.
 */
export class OfficeUiGallery extends OfficeElement {
	static pickEvent = 'office-gallery-pick';
	static previewEvent = 'office-gallery-preview';
	static previewEndEvent = 'office-gallery-preview-end';
	/** Hover time before the first preview shows, so sweeping across the strip does not flash. */
	static previewDelayMs = 120;
	/** Grace before a preview ends, so moving between tiles does not flash the original. */
	static previewEndDelayMs = 60;
	static triggerAttribute = 'data-gallery';
	static popupAttribute = 'data-gallery-popup';
	static itemAttribute = 'data-gallery-item';
	static compactAttribute: string | null = null;
	static override styles = controlStyles(HOST_CSS);
	/** `mode`, `chevron-only`, `icon` and `compact` are read from the attributes. */
	static override get observedAttributes(): string[] {
		return [...super.observedAttributes, 'mode', 'chevron-only', 'icon', 'compact'];
	}
	private model: OfficeGalleryState | undefined;
	private locked = false;
	private opened = false;
	private previewing: string | undefined;
	private previewStartTimer: ReturnType<typeof setTimeout> | undefined;
	private previewEndTimer: ReturnType<typeof setTimeout> | undefined;
	private menu: { itemId: string; x: number; y: number } | undefined;

	override attributeChangedCallback(name: string, old: string | null, value: string | null): void {
		super.attributeChangedCallback(name, old, value);
		if (name === 'mode' && value === 'panel') this.close();
		this.requestUpdate();
	}

	get state(): OfficeGalleryState | undefined {
		return this.model;
	}
	set state(value: OfficeGalleryState | null | undefined) {
		const next = value ?? undefined;
		// A host that sets its state on every render of its own hands over an equal one most of the
		// time: drawing it again would rebuild every tile for nothing and make them flash.
		const json = next ? JSON.stringify(next) : '';
		if (this.stateJson === json && this.hasUpdated) return;
		this.stateJson = json;
		this.model = next;
		this.pruneParsed(next);
		this.requestUpdate('state');
	}
	private stateJson = '';
	/** The parsed preview of each tile, kept while its markup is unchanged so Lit leaves it alone. */
	private parsed = new Map<string, { markup: string; node: Element | null }>();
	private parsedPreview(slot: string, item: OfficeGalleryItem): Element | null {
		if (!item.preview) return null;
		const key = `${slot}::${item.id}`;
		const cached = this.parsed.get(key);
		if (cached?.markup === item.preview) return cached.node;
		const node = parseSvgPreview(this.ownerDocument, item.preview);
		this.parsed.set(key, { markup: item.preview, node });
		return node;
	}
	/** Forget previews of tiles the new state no longer has. */
	private pruneParsed(state: OfficeGalleryState | undefined): void {
		const keep = new Set<string>();
		for (const item of state?.inline ?? []) keep.add(`inline::${item.id}`);
		for (const section of state?.sections ?? [])
			for (const item of section.items) keep.add(`${section.title ?? ''}::${item.id}`);
		for (const key of this.parsed.keys()) if (!keep.has(key)) this.parsed.delete(key);
	}

	get disabled(): boolean {
		return this.locked;
	}
	set disabled(value: unknown) {
		this.locked = present(value);
		this.requestUpdate('disabled');
	}

	get open(): boolean {
		return this.opened;
	}
	set open(value: unknown) {
		const next = present(value) && !this.unavailable() && this.getAttribute('mode') !== 'panel';
		if (next === this.opened) return;
		this.opened = next;
		this.cleanup();
		this.requestUpdate('open');
		if (next && this.isConnected) {
			const doc = this.ownerDocument;
			doc.addEventListener('pointerdown', this.onOutside, true);
			doc.addEventListener('keydown', this.onEscape);
			doc.defaultView?.addEventListener('resize', this.position);
			doc.addEventListener('scroll', this.position, true);
			this.position();
		}
	}

	close(): void {
		this.endPreview(true);
		this.closeMenu();
		this.open = false;
	}

	/** Close the popup and menu but leave a running preview for the pick that follows to settle. */
	private dismiss(): void {
		clearTimeout(this.previewStartTimer);
		this.previewStartTimer = undefined;
		this.closeMenu();
		this.open = false;
	}

	// ---- Preview and context menu ---------------------------------------------------------

	private canPreview(): boolean {
		return (
			this.model?.preview === true &&
			!this.model.command &&
			!this.unavailable() &&
			this.getAttribute('mode') !== 'panel'
		);
	}

	private startPreview(itemId: string): void {
		if (!this.canPreview() || this.menu) return;
		clearTimeout(this.previewEndTimer);
		if (this.previewing === itemId) return;
		clearTimeout(this.previewStartTimer);
		const show = (): void => {
			const state = this.model;
			if (!state || !this.canPreview()) return;
			this.previewing = itemId;
			this.fire((this.constructor as unknown as Statics).previewEvent, {
				gallery: state.id,
				itemId,
			});
		};
		const statics = this.constructor as unknown as Statics & { previewDelayMs: number };
		// Moving from one previewed tile to the next is immediate; the first one waits for a pause.
		if (this.previewing !== undefined) show();
		else this.previewStartTimer = setTimeout(show, statics.previewDelayMs);
	}

	private endPreview(immediate = false): void {
		clearTimeout(this.previewStartTimer);
		this.previewStartTimer = undefined;
		clearTimeout(this.previewEndTimer);
		this.previewEndTimer = undefined;
		if (this.previewing === undefined) return;
		const end = (): void => {
			this.previewEndTimer = undefined;
			const state = this.model;
			if (this.previewing === undefined) return;
			this.previewing = undefined;
			this.fire((this.constructor as unknown as Statics).previewEndEvent, {
				gallery: state?.id ?? '',
			});
		};
		const statics = this.constructor as unknown as Statics & { previewEndDelayMs: number };
		if (immediate) end();
		else this.previewEndTimer = setTimeout(end, statics.previewEndDelayMs);
	}

	private findItem(itemId: string): OfficeGalleryItem | undefined {
		for (const section of this.model?.sections ?? []) {
			const found = section.items.find((item) => item.id === itemId);
			if (found) return found;
		}
		return this.model?.inline?.find((item) => item.id === itemId);
	}

	private openMenu(itemId: string, x: number, y: number): void {
		if (this.unavailable() || !this.findItem(itemId)?.actions?.length) return;
		this.endPreview(true);
		this.menu = { itemId, x, y };
		this.ownerDocument.addEventListener('pointerdown', this.onMenuOutside, true);
		this.requestUpdate();
	}

	private closeMenu(): void {
		if (!this.menu) return;
		this.menu = undefined;
		this.ownerDocument.removeEventListener('pointerdown', this.onMenuOutside, true);
		this.requestUpdate();
	}

	private readonly onMenuOutside = (event: Event): void => {
		if (!event.composedPath().some((node) => (node as Element).classList?.contains('tile-menu')))
			this.closeMenu();
	};

	private onTileContextMenu(event: MouseEvent, itemId: string): void {
		if (!this.findItem(itemId)?.actions?.length) return;
		event.preventDefault();
		event.stopPropagation();
		this.openMenu(itemId, event.clientX, event.clientY);
	}

	private onMenuKey(event: KeyboardEvent): void {
		const items = Array.from(
			(event.currentTarget as HTMLElement).querySelectorAll<HTMLButtonElement>('button'),
		);
		const index = items.indexOf(event.target as HTMLButtonElement);
		if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			const itemId = this.menu?.itemId;
			this.closeMenu();
			const statics = this.constructor as unknown as Statics;
			this.querySelector<HTMLElement>(`[${statics.itemAttribute}="${itemId}"]`)?.focus();
		} else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			event.preventDefault();
			event.stopPropagation();
			const step = event.key === 'ArrowDown' ? 1 : -1;
			items[(index + step + items.length) % items.length]?.focus();
		} else if (event.key === 'Tab') {
			this.closeMenu();
		}
	}

	private menuTemplate(): TemplateResult | string {
		const menu = this.menu;
		const item = menu ? this.findItem(menu.itemId) : undefined;
		if (!menu || !item?.actions?.length) return '';
		return html`<div
			class="tile-menu"
			role="menu"
			aria-label=${item.label}
			style="left:${menu.x}px;top:${menu.y}px"
			@keydown=${this.onMenuKey}
		>
			${item.actions.map(
				(action) =>
					html`<button
						type="button"
						role="menuitem"
						class="tile-menu-item"
						data-gallery-action=${action.id}
						@click=${() => this.pick(menu.itemId, action.id)}
						>${action.label}</button
					>`,
			)}
		</div>`;
	}

	get trigger(): HTMLButtonElement {
		this.ensureRendered();
		return this.querySelector<HTMLButtonElement>('.trigger')!;
	}
	/** The open popup; while closed it is not in the DOM, so this is a detached hidden placeholder. */
	get popup(): HTMLElement {
		this.ensureRendered();
		const open = this.querySelector<HTMLElement>('.popup');
		if (open) return open;
		// Keep the host's popup hook on the placeholder so callers can read it while closed.
		const attribute = (this.constructor as unknown as Statics).popupAttribute;
		const id = this.state?.id;
		if (id === undefined) this.closedPopup.removeAttribute(attribute);
		else this.closedPopup.setAttribute(attribute, id);
		return this.closedPopup;
	}
	private readonly closedPopup = Object.assign(this.ownerDocument.createElement('div'), {
		className: 'popup',
		hidden: true,
	});

	/** The trigger glyph: `command.icon` for a command, else the `icon` attribute. */
	triggerIcon(): Element | null {
		const name = this.model?.command?.icon ?? this.getAttribute('icon');
		if (!name) return null;
		// One node for the gallery's life: a new one each render would be swapped into the page.
		const svg = (this.iconNode ??= createIconSvg(this.ownerDocument));
		return paintIcon(svg, name) ? svg : null;
	}
	private iconNode: SVGSVGElement | undefined;
	private emptyIcon: SVGSVGElement | undefined;

	override connectedCallback(): void {
		this.setAttribute('data-office-gallery', '');
		attachGalleryStyles(this);
		super.connectedCallback();
		this.markHost();
	}

	/** Host attributes (never written before the element is connected). */
	private markHost(): void {
		if (!this.isConnected) return;
		this.toggleAttribute('data-command', Boolean(this.model?.command));
		this.toggleAttribute('data-command-large', Boolean(this.model?.command?.large));
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.close();
		this.cleanup();
		clearTimeout(this.previewStartTimer);
		clearTimeout(this.previewEndTimer);
	}

	private unavailable(): boolean {
		const state = this.model;
		return (
			this.locked ||
			!state ||
			state.disabled === true ||
			!state.sections.some((section) => section.items.length > 0)
		);
	}

	private pick(itemId: string, action?: string): void {
		const state = this.model;
		if (
			this.unavailable() ||
			!state ||
			!(
				state.sections.some((section) => section.items.some((item) => item.id === itemId)) ||
				state.footer?.some((entry) => entry.id === itemId)
			)
		)
			return;
		// The pick goes out while the preview still stands, so the host settles on the state it
		// previewed from; the preview ends right after.
		this.dismiss();
		if (this.getAttribute('mode') !== 'panel') this.trigger.focus();
		this.fire(
			(this.constructor as unknown as Statics).pickEvent,
			action === undefined ? { gallery: state.id, itemId } : { gallery: state.id, itemId, action },
		);
		this.endPreview(true);
	}

	private onTrigger(): void {
		if (this.model?.command) {
			this.pick(this.model.sections[0]?.items[0]?.id ?? '');
			return;
		}
		this.open = !this.open;
	}

	private onKey(event: KeyboardEvent): void {
		const panel = this.getAttribute('mode') === 'panel';
		if (panel || this.open) {
			const target = event.target as HTMLElement;
			const tile = target.closest<HTMLButtonElement>('.tile');
			const section = tile?.closest('.section');
			const items = Array.from(
				(panel ? section : this.popup)?.querySelectorAll<HTMLButtonElement>(
					'.tile:not(:disabled)',
				) ?? [],
			);
			const index = tile ? items.indexOf(tile) : -1;
			const columns =
				this.model?.sections[Array.from(this.querySelectorAll('.section')).indexOf(section!)]
					?.columns ?? 1;
			const delta = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: columns, ArrowUp: -columns }[
				event.key
			];
			const next =
				event.key === 'Home'
					? 0
					: event.key === 'End'
						? items.length - 1
						: delta === undefined
							? -1
							: (index + delta + items.length) % items.length;
			if (index >= 0 && next >= 0) {
				event.preventDefault();
				event.stopPropagation();
				items[next]?.focus();
				// A compact panel applies its pick at once, so browsing must not apply each tile passed.
				if (panel && !this.hasAttribute('compact')) items[next]?.click();
				return;
			}
			if (panel) return;
		}
		const trigger = this.trigger;
		if (event.key === 'Escape' && this.open) {
			event.stopPropagation();
			event.preventDefault();
			this.close();
			trigger.focus();
		}
		if (event.key === 'ArrowDown' && event.target === trigger) {
			event.stopPropagation();
			event.preventDefault();
			this.open = true;
			this.popup.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
		}
		if (event.key === 'Enter' || event.key === ' ') event.stopPropagation();
	}

	/** The Menu key and Shift+F10 open a tile's context menu where the tile is. */
	private onTileKey(event: KeyboardEvent, itemId: string): void {
		if (event.key !== 'ContextMenu' && !(event.key === 'F10' && event.shiftKey)) return;
		if (!this.findItem(itemId)?.actions?.length) return;
		event.preventDefault();
		event.stopPropagation();
		const box = (event.currentTarget as HTMLElement).getBoundingClientRect();
		this.openMenu(itemId, box.left, box.bottom);
	}

	private readonly onOutside = (event: Event): void => {
		if (!event.composedPath().includes(this)) this.close();
	};

	private readonly onEscape = (event: KeyboardEvent): void => {
		if (event.key !== 'Escape') return;
		event.stopPropagation();
		this.close();
		this.trigger.focus();
	};

	private readonly position = (): void => {
		const window = this.ownerDocument.defaultView;
		const popup = this.querySelector<HTMLElement>('.popup');
		if (!window || !popup) return;
		const anchor = this.trigger.getBoundingClientRect();
		const box = popup.getBoundingClientRect();
		// Viewport geometry, so it is set inline; the margin keeps the popup off the window edge.
		const margin = 8;
		const left = Math.max(margin, Math.min(anchor.left, window.innerWidth - box.width - margin));
		const top = Math.max(
			margin,
			Math.min(anchor.bottom + 4, window.innerHeight - box.height - margin),
		);
		popup.style.left = `${left}px`;
		popup.style.top = `${top}px`;
		// An ancestor with `contain`, `transform` or `filter` becomes the fixed popup's containing
		// block and offsets it from the viewport; measure where it landed and take that back out.
		const placed = popup.getBoundingClientRect();
		if (placed.width > 0 && (placed.left !== left || placed.top !== top)) {
			popup.style.left = `${2 * left - placed.left}px`;
			popup.style.top = `${2 * top - placed.top}px`;
		}
	};

	private cleanup(): void {
		const doc = this.ownerDocument;
		doc.removeEventListener('pointerdown', this.onOutside, true);
		doc.removeEventListener('keydown', this.onEscape);
		doc.defaultView?.removeEventListener('resize', this.position);
		doc.removeEventListener('scroll', this.position, true);
	}

	// ---- Light-DOM template ---------------------------------------------------------------

	private tile(
		slot: string,
		item: OfficeGalleryItem,
		width: number,
		height: number,
		disabled: boolean,
	) {
		const { itemAttribute, compactAttribute } = this.constructor as unknown as Statics;
		const preview = this.parsedPreview(slot, item);
		// Tile size is document data (the preview's own aspect), so it is set inline.
		return staticHtml`<button
			type="button"
			class="tile"
			${unsafeStatic(compactAttribute ?? '')}
			${unsafeStatic(itemAttribute)}=${item.id}
			?disabled=${disabled}
			aria-pressed=${String(item.applied === true)}
			aria-label=${item.label}
			title=${item.label}
			style="width:${width + 6}px;height:${height + 6}px;--office-gallery-preview-width:${width}px;--office-gallery-preview-height:${height}px"
			@click=${() => this.pick(item.id)}
			@pointerenter=${(event: PointerEvent) => {
				if (event.pointerType !== 'touch') this.startPreview(item.id);
			}}
			@pointerleave=${() => this.endPreview()}
			@focus=${(event: FocusEvent) => {
				if ((event.currentTarget as HTMLElement).matches(':focus-visible'))
					this.startPreview(item.id);
			}}
			@blur=${() => this.endPreview()}
			@contextmenu=${(event: MouseEvent) => this.onTileContextMenu(event, item.id)}
			@keydown=${(event: KeyboardEvent) => this.onTileKey(event, item.id)}
		>${preview}${this.getAttribute('mode') === 'panel' && !this.hasAttribute('compact') ? html`<span class="tile-label">${item.label}</span>` : ''}</button>`;
	}

	private lightTemplate(): TemplateResult {
		const statics = this.constructor as unknown as Statics;
		const state = this.model;
		const disabled = this.unavailable();
		if (this.getAttribute('mode') === 'panel')
			return html`<div
				class="gallery-panel"
				role="group"
				aria-label=${state?.label ?? ''}
				@keydown=${this.onKey}
				>${(state?.sections ?? []).map((section) => this.section(section, disabled))}</div
			>`;
		const inline = this.getAttribute('mode') === 'inline' && !this.hasAttribute('chevron-only');
		const chevronOnly = this.hasAttribute('chevron-only');
		const command = state?.command;
		const title = state?.label ?? '';
		const first = state?.sections[0];
		const inlineItems = state?.inline ?? first?.items ?? [];
		const triggerTitle = command
			? (command.hint ?? title)
			: inline
				? (state?.moreLabel ?? `More ${title}`)
				: title;
		const icon = this.triggerIcon() ?? (this.emptyIcon ??= createIconSvg(this.ownerDocument));
		const trigger = staticHtml`<button
			type="button"
			class="trigger${command ? ' command' : ''}${command?.large ? ' command-large' : ''}"
			${unsafeStatic(statics.compactAttribute ?? '')}
			${unsafeStatic(statics.triggerAttribute)}=${ifDefined(state?.id)}
			title=${triggerTitle}
			aria-label=${command ? title : triggerTitle}
			aria-haspopup=${ifDefined(command ? undefined : 'dialog')}
			aria-expanded=${ifDefined(command || !state ? undefined : String(this.opened))}
			?disabled=${!state || disabled}
			@click=${this.onTrigger}
		>${command || (!inline && !chevronOnly) ? html`${icon}<span>${title}</span>` : ''}${command ? '' : html`<span aria-hidden="true">⌄</span>`}</button>`;
		const popup = staticHtml`<div
			class="popup"
			role="dialog"
			aria-label=${title}
			${unsafeStatic(statics.popupAttribute)}=${ifDefined(state?.id)}
		>${(state?.sections ?? []).map((section) => this.section(section, disabled))}${this.footer(disabled)}</div>`;
		return html`<span class="gallery-view" @keydown=${this.onKey}>
			<div class="strip" ?hidden=${!inline} style=${inline ? '' : 'display:none'}>
				${
					inline
						? repeat(
								inlineItems,
								(item) => item.id,
								(item) =>
									this.tile(
										'inline',
										item,
										first?.tileWidth ?? 40,
										first?.tileHeight ?? 30,
										disabled,
									),
							)
						: ''
				}
			</div>
			${trigger}${this.opened ? popup : ''}${this.menuTemplate()}
		</span>`;
	}

	private footer(disabled: boolean) {
		const entries = this.model?.footer ?? [];
		if (entries.length === 0) return '';
		return html`<div class="footer" role="group">
			${entries.map(
				(entry) =>
					html`<button
						type="button"
						class="footer-item"
						data-gallery-footer=${entry.id}
						?disabled=${disabled}
						@click=${() => this.pick(entry.id)}
						>${entry.label}</button
					>`,
			)}
		</div>`;
	}

	private section(section: OfficeGallerySection, disabled: boolean) {
		return html`<section class="section">
			${section.title ? html`<div class="heading">${section.title}</div>` : ''}
			<div class="grid" style="grid-template-columns:repeat(${section.columns}, minmax(0, 1fr))">
				${repeat(
					section.items,
					(item) => item.id,
					(item) =>
						this.tile(section.title ?? '', item, section.tileWidth, section.tileHeight, disabled),
				)}
			</div>
		</section>`;
	}

	/** The tiles are light-DOM children of the host, shown through this slot. */
	protected override render() {
		return html`<slot></slot>`;
	}

	protected override updated(): void {
		const state = this.model;
		if (state && this.unavailable() && this.opened) this.close();
		// The tiles live in the light DOM, rendered into the host itself (not from a constructor).
		if (this.detachedFirstRender) return;
		render(this.lightTemplate(), this, { host: this });
		this.markHost();
		if (this.opened) this.position();
		const menu = this.querySelector<HTMLElement>('.tile-menu');
		if (menu && !menu.contains(this.ownerDocument.activeElement))
			menu.querySelector<HTMLElement>('button')?.focus();
	}
}

export const defineGallery = definer('office-ui-gallery', () => OfficeUiGallery);

export { parseSvgPreview } from './safe-svg';
