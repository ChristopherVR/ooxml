import { html, render, type TemplateResult } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { repeat } from 'lit/directives/repeat.js';
import { html as staticHtml, unsafeStatic } from 'lit/static-html.js';
import { OfficeElement, controlStyles } from '../base.js';
import { createIconSvg, paintIcon } from '../icons.js';
import { definer, present } from '../registry.js';
import { attachGalleryStyles } from './gallery-styles.js';
import { parseSvgPreview } from './safe-svg.js';

export type OfficeGalleryPickEvent = CustomEvent<{ gallery: string; itemId: string }>;

/** One tile. `preview` is SVG markup, parsed as an SVG document with executable parts removed. */
export interface OfficeGalleryItem {
	id: string;
	label: string;
	applied?: boolean | undefined;
	preview?: string | undefined;
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
}

type Statics = {
	pickEvent: string;
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
 * and closes on Escape, an outside press or a pick. ArrowDown on the trigger opens it and focuses
 * the first tile. A pick emits `office-gallery-pick` `{ gallery, itemId }`.
 *
 * Tiles render in the light DOM beside the host's children (Lit renders the template into the
 * host itself) so products can query them (`data-gallery-item`); the rules join the host's root
 * once. Static `pickEvent`, `triggerAttribute`, `popupAttribute`, `itemAttribute` and
 * `compactAttribute` keep a product's published names; override `triggerIcon()` to draw a product
 * glyph.
 */
export class OfficeUiGallery extends OfficeElement {
	static pickEvent = 'office-gallery-pick';
	static triggerAttribute = 'data-gallery';
	static popupAttribute = 'data-gallery-popup';
	static itemAttribute = 'data-gallery-item';
	static compactAttribute: string | null = null;
	static override styles = controlStyles(HOST_CSS);
	/** `mode`, `chevron-only` and `icon` are read from the attributes, as before. */
	static override get observedAttributes(): string[] {
		return [...super.observedAttributes, 'mode', 'chevron-only', 'icon'];
	}
	private model: OfficeGalleryState | undefined;
	private locked = false;
	private opened = false;

	override attributeChangedCallback(name: string, old: string | null, value: string | null): void {
		super.attributeChangedCallback(name, old, value);
		this.requestUpdate();
	}

	get state(): OfficeGalleryState | undefined {
		return this.model;
	}
	set state(value: OfficeGalleryState | null | undefined) {
		this.model = value ?? undefined;
		this.requestUpdate('state');
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
		const next = present(value) && !this.unavailable();
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
		this.open = false;
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
		const svg = createIconSvg(this.ownerDocument);
		return paintIcon(svg, name) ? svg : null;
	}

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

	private pick(itemId: string): void {
		const state = this.model;
		if (
			this.unavailable() ||
			!state?.sections.some((section) => section.items.some((item) => item.id === itemId))
		)
			return;
		this.close();
		this.trigger.focus();
		this.fire((this.constructor as unknown as Statics).pickEvent, { gallery: state.id, itemId });
	}

	private onTrigger(): void {
		if (this.model?.command) {
			this.pick(this.model.sections[0]?.items[0]?.id ?? '');
			return;
		}
		this.open = !this.open;
	}

	private onKey(event: KeyboardEvent): void {
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
		popup.style.left = `${Math.max(margin, Math.min(anchor.left, window.innerWidth - box.width - margin))}px`;
		popup.style.top = `${Math.max(margin, Math.min(anchor.bottom + 4, window.innerHeight - box.height - margin))}px`;
	};

	private cleanup(): void {
		const doc = this.ownerDocument;
		doc.removeEventListener('pointerdown', this.onOutside, true);
		doc.removeEventListener('keydown', this.onEscape);
		doc.defaultView?.removeEventListener('resize', this.position);
		doc.removeEventListener('scroll', this.position, true);
	}

	// ---- Light-DOM template ---------------------------------------------------------------

	private tile(item: OfficeGalleryItem, width: number, height: number, disabled: boolean) {
		const { itemAttribute, compactAttribute } = this.constructor as unknown as Statics;
		const preview = item.preview ? parseSvgPreview(this.ownerDocument, item.preview) : null;
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
			style="width:${width + 6}px;height:${height + 6}px"
			@click=${() => this.pick(item.id)}
		>${preview}</button>`;
	}

	private lightTemplate(): TemplateResult {
		const statics = this.constructor as unknown as Statics;
		const state = this.model;
		const disabled = this.unavailable();
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
		const icon = this.triggerIcon() ?? createIconSvg(this.ownerDocument);
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
		>${(state?.sections ?? []).map((section) => this.section(section, disabled))}</div>`;
		return html`<span class="gallery-view" @keydown=${this.onKey}>
			<div class="strip" ?hidden=${!inline} style=${inline ? '' : 'display:none'}>
				${
					inline
						? repeat(
								inlineItems,
								(item) => item.id,
								(item) =>
									this.tile(item, first?.tileWidth ?? 40, first?.tileHeight ?? 30, disabled),
							)
						: ''
				}
			</div>
			${trigger}${this.opened ? popup : ''}
		</span>`;
	}

	private section(section: OfficeGallerySection, disabled: boolean) {
		return html`<section class="section">
			${section.title ? html`<div class="heading">${section.title}</div>` : ''}
			<div class="grid" style="grid-template-columns:repeat(${section.columns}, minmax(0, 1fr))">
				${repeat(
					section.items,
					(item) => item.id,
					(item) => this.tile(item, section.tileWidth, section.tileHeight, disabled),
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
	}
}

export const defineGallery = definer('office-ui-gallery', () => OfficeUiGallery);

export { parseSvgPreview } from './safe-svg.js';
