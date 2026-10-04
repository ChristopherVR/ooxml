import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit, present } from './registry.js';
import { attachStyles, controlCss } from './styles.js';
import { attachGalleryStyles } from './gallery-styles.js';
import { createGalleryView, type GalleryHooks, type OfficeGalleryState } from './gallery-view.js';

export type OfficeGalleryPickEvent = CustomEvent<{ gallery: string; itemId: string }>;

const HOST_CSS = `
:host { display: inline-flex; position: relative; flex: none; }
`;

type Statics = GalleryHooks & { pickEvent: string };

/**
 * `<office-ui-gallery>`: Office's ribbon gallery (Shape Styles, Themes, Table Styles). Set
 * `state` (`OfficeGalleryState`, translated). `mode="inline"` shows a strip of tiles beside a
 * More button; otherwise the trigger is a labelled dropdown (`chevron-only` keeps just the
 * chevron, `icon` names its glyph). The popup opens below the trigger, keeps inside the window
 * and closes on Escape, an outside press or a pick. ArrowDown on the trigger opens it and focuses
 * the first tile. A pick emits `office-gallery-pick` `{ gallery, itemId }`.
 *
 * Tiles render in the light DOM beside the host's children so products can query them
 * (`data-gallery-item`); the rules join the host's root once. Static `pickEvent`,
 * `triggerAttribute`, `popupAttribute`, `itemAttribute` and `compactAttribute` keep a product's
 * published names; override `triggerIcon()` to draw a product glyph.
 */
export const defineGallery = definer('office-ui-gallery', () => {
	class OfficeUiGallery extends HTMLElement {
		static pickEvent = 'office-gallery-pick';
		static triggerAttribute = 'data-gallery';
		static popupAttribute = 'data-gallery-popup';
		static itemAttribute = 'data-gallery-item';
		static compactAttribute: string | null = null;
		static observedAttributes = ['mode', 'chevron-only', 'icon'];
		#state: OfficeGalleryState | undefined;
		#locked = false;
		#opened = false;
		readonly #view;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(HOST_CSS));
			root.append(this.ownerDocument.createElement('slot'));
			const statics = this.constructor as unknown as Statics;
			this.#view = createGalleryView(
				this.ownerDocument,
				statics,
				(itemId) => this.#pick(itemId),
				() => {
					if (this.#state?.command) {
						this.#pick(this.#state.sections[0]?.items[0]?.id ?? '');
						return;
					}
					this.open = !this.open;
				},
			);
			this.#view.root.addEventListener('keydown', (event) => {
				if (event.key === 'Escape' && this.open) {
					event.stopPropagation();
					event.preventDefault();
					this.close();
					this.#view.trigger.focus();
				}
				if (event.key === 'ArrowDown' && event.target === this.#view.trigger) {
					event.stopPropagation();
					event.preventDefault();
					this.open = true;
					this.#view.popup.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
				}
				if (event.key === 'Enter' || event.key === ' ') event.stopPropagation();
			});
		}
		get state(): OfficeGalleryState | undefined {
			return this.#state;
		}
		set state(value: OfficeGalleryState | null | undefined) {
			this.#state = value ?? undefined;
			this.#paint();
		}
		get trigger(): HTMLButtonElement {
			return this.#view.trigger;
		}
		get popup(): HTMLElement {
			return this.#view.popup;
		}
		get disabled(): boolean {
			return this.#locked;
		}
		set disabled(value: unknown) {
			this.#locked = present(value);
			this.#paint();
		}
		get open(): boolean {
			return this.#opened;
		}
		set open(value: unknown) {
			const next = present(value) && !this.#unavailable();
			if (next === this.#opened) return;
			this.#opened = next;
			this.#view.setOpen(next);
			this.#paint();
			this.#view.trigger.setAttribute('aria-expanded', String(next));
			this.#cleanup();
			if (next && this.isConnected) {
				const doc = this.ownerDocument;
				doc.addEventListener('pointerdown', this.#outside, true);
				doc.addEventListener('keydown', this.#escape);
				doc.defaultView?.addEventListener('resize', this.#position);
				doc.addEventListener('scroll', this.#position, true);
				this.#position();
			}
		}
		close(): void {
			this.open = false;
		}
		/** The trigger glyph: `command.icon` for a command, else the `icon` attribute. */
		triggerIcon(): Element | null {
			const name = this.#state?.command?.icon ?? this.getAttribute('icon');
			if (!name) return null;
			const svg = createIconSvg(this.ownerDocument);
			return paintIcon(svg, name) ? svg : null;
		}
		connectedCallback(): void {
			this.setAttribute('data-office-gallery', '');
			attachGalleryStyles(this);
			this.append(this.#view.root);
			this.#paint();
		}
		disconnectedCallback(): void {
			this.close();
			this.#cleanup();
		}
		attributeChangedCallback(): void {
			this.#paint();
		}
		#unavailable(): boolean {
			const state = this.#state;
			return (
				this.#locked ||
				!state ||
				state.disabled === true ||
				!state.sections.some((section) => section.items.length > 0)
			);
		}
		#paint(): void {
			const state = this.#state;
			if (!state) {
				this.close();
				this.#view.clear();
				this.#view.trigger.disabled = true;
				return;
			}
			const doc = this.ownerDocument;
			const hadFocus = this.contains(doc.activeElement);
			if (this.isConnected && this.#view.root.parentElement !== this) this.append(this.#view.root);
			const itemAttribute = (this.constructor as unknown as Statics).itemAttribute;
			const focusedId = doc.activeElement?.getAttribute(itemAttribute);
			const focusInPopup = this.#view.popup.contains(doc.activeElement);
			const chevronOnly = this.hasAttribute('chevron-only');
			this.#view.paint({
				state,
				disabled: this.#unavailable(),
				inline: this.getAttribute('mode') === 'inline' && !chevronOnly,
				chevronOnly,
				open: this.#opened,
				icon: this.triggerIcon(),
			});
			if (state.command) this.#view.trigger.removeAttribute('aria-expanded');
			else this.#view.trigger.setAttribute('aria-expanded', String(this.#opened));
			if (this.isConnected) {
				this.toggleAttribute('data-command', Boolean(state.command));
				this.toggleAttribute('data-command-large', Boolean(state.command?.large));
			}
			if (this.#unavailable()) this.close();
			if (this.#opened) this.#position();
			if (focusedId && hadFocus)
				[
					...(focusInPopup ? this.#view.popup : this).querySelectorAll<HTMLButtonElement>(
						`[${itemAttribute}]`,
					),
				]
					.find((node) => node.getAttribute(itemAttribute) === focusedId)
					?.focus();
		}
		#pick(itemId: string): void {
			const state = this.#state;
			if (
				this.#unavailable() ||
				!state?.sections.some((section) => section.items.some((item) => item.id === itemId))
			)
				return;
			this.close();
			this.#view.trigger.focus();
			emit(this, (this.constructor as unknown as Statics).pickEvent, {
				gallery: state.id,
				itemId,
			});
		}
		readonly #outside = (event: Event): void => {
			if (!event.composedPath().includes(this)) this.close();
		};
		readonly #escape = (event: KeyboardEvent): void => {
			if (event.key === 'Escape') {
				event.stopPropagation();
				this.close();
				this.#view.trigger.focus();
			}
		};
		readonly #position = (): void => {
			const window = this.ownerDocument.defaultView;
			if (!window) return;
			const popup = this.#view.popup;
			const anchor = this.#view.trigger.getBoundingClientRect();
			const box = popup.getBoundingClientRect();
			// Viewport geometry, so it is set inline; 8px keeps the popup off the window edge.
			popup.style.left = `${Math.max(8, Math.min(anchor.left, window.innerWidth - box.width - 8))}px`;
			popup.style.top = `${Math.max(8, Math.min(anchor.bottom + 4, window.innerHeight - box.height - 8))}px`;
		};
		#cleanup(): void {
			const doc = this.ownerDocument;
			doc.removeEventListener('pointerdown', this.#outside, true);
			doc.removeEventListener('keydown', this.#escape);
			doc.defaultView?.removeEventListener('resize', this.#position);
			doc.removeEventListener('scroll', this.#position, true);
		}
	}
	return OfficeUiGallery;
});

export type {
	OfficeGalleryItem,
	OfficeGallerySection,
	OfficeGalleryState,
} from './gallery-view.js';
export { parseSvgPreview } from './safe-svg.js';
