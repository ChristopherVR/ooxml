import { LitElement, html, type PropertyValues } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeElement, controlStyles, flag } from '../base.js';
import { definer, present } from '../registry.js';
import {
	collectSelectChoices,
	nextSelectActive,
	pageSelectActive,
	positionSelectMenu,
	SELECT_OPTION_ATTRIBUTES,
	type SelectChoice,
} from './select-menu.js';
import css from './select.css?raw';

/** A programmatic choice; `<option>` elements satisfy the same shape. */
export interface OfficeSelectOption {
	value: string;
	label: string;
	disabled?: boolean;
}

let uid = 0;

/**
 * Single-select combobox. Moved from pptx-viewer `pptx-ui-select` (the declarative, most complete
 * implementation). Choices come from `<option>` / `<optgroup>` children (with `data-display-label`,
 * `data-description` and font previews through `style="font-family"`), or from the `options`
 * property when there are none. `value`, `disabled`, `selectedIndex`, `aria-label`;
 * `variant="ribbon-font"` (compact ribbon combo that accepts unlisted values, such as a typed font
 * size) or `"ribbon-icon"`; `data-font-picker="family|size"`; an `icon` slot and a `custom` slot at
 * the top of the popup. Keyboard: arrows, Home, End, PageUp/PageDown (8 options), typeahead,
 * Enter/Space, Escape, Tab; disabled and hidden choices are skipped. The popup is a top-layer
 * `popover`, flipped above the trigger when there is more room. `input` then `change` bubble on a
 * user choice; setting a property is silent. Form-associated.
 */
export class OfficeUiSelect extends OfficeElement {
	static formAssociated = true;
	static override shadowRootOptions = { ...LitElement.shadowRootOptions, delegatesFocus: true };
	static override styles = controlStyles(css);
	static override properties = {
		// `value` and `options` are derived from the `<option>` children, so they have accessors.
		value: { type: String, reflect: true, noAccessor: true },
		options: { attribute: false, noAccessor: true },
		disabled: flag,
		open: flag,
		active: { state: true },
		choices: { state: true },
		opened: { state: true },
	};
	declare disabled: boolean;
	declare open: boolean;
	declare active: number;
	declare choices: SelectChoice[];
	declare opened: boolean;

	static override get observedAttributes(): string[] {
		return [...super.observedAttributes, 'aria-label', 'variant', 'data-font-picker'];
	}

	private readonly internals: ElementInternals | undefined;
	private readonly observer: MutationObserver | undefined;
	private readonly menuId = `office-ui-select-options-${++uid}`;
	private requested: string | null = null;
	private data: OfficeSelectOption[] = [];
	private refreshFrame = 0;
	private search = '';
	private searchTimer = 0;
	private defaultValue = '';

	constructor() {
		super();
		this.disabled = false;
		this.open = false;
		this.active = -1;
		this.choices = [];
		this.opened = false;
		try {
			this.internals = this.attachInternals();
		} catch {
			this.internals = undefined;
		}
		const Observer = this.ownerDocument.defaultView?.MutationObserver;
		this.observer = Observer
			? new Observer(() => {
					// Let frameworks finish updating their option nodes before touching the shadow DOM.
					const view = this.ownerDocument.defaultView;
					if (this.refreshFrame || !view) return;
					this.refreshFrame = view.requestAnimationFrame(() => {
						this.refreshFrame = 0;
						this.refresh();
					});
				})
			: undefined;
	}

	/** The `<option>` children, or the programmatic options when there are none. */
	get options(): OfficeSelectOption[] {
		const children = [...this.querySelectorAll('option')];
		return children.length ? children : this.data;
	}
	set options(next: OfficeSelectOption[]) {
		this.data = [...next];
		this.refresh();
	}

	get value(): string {
		const requested = this.requested;
		const choices = this.choices.length ? this.choices : this.collect();
		if (requested !== null && this.getAttribute('variant') === 'ribbon-font') return requested;
		if (requested !== null && (!choices.length || choices.some((c) => c.value === requested)))
			return requested;
		return (
			[...this.querySelectorAll('option')].find((option) => option.selected)?.value ??
			choices.find((choice) => !choice.disabled && !choice.hidden)?.value ??
			choices[0]?.value ??
			requested ??
			''
		);
	}
	set value(next: string) {
		this.requested = String(next ?? '');
		// Always redraw: the choices may have changed even when the value reads the same.
		this.requestUpdate('value', undefined);
	}

	get selectedIndex(): number {
		return this.options.findIndex((option) => option.value === this.value);
	}
	set selectedIndex(index: number) {
		this.value = this.options[index]?.value ?? '';
	}

	private get menu(): HTMLElement | null {
		return this.renderRoot.querySelector('.menu');
	}
	private get trigger(): HTMLButtonElement | null {
		return this.renderRoot.querySelector('button');
	}

	override focus(options?: FocusOptions): void {
		this.trigger?.focus(options);
	}

	formResetCallback(): void {
		this.value = this.defaultValue;
	}

	formDisabledCallback(disabled: boolean): void {
		this.disabled = disabled;
	}

	override connectedCallback(): void {
		this.defaultValue = this.getAttribute('value') ?? '';
		this.observer?.observe(this, {
			childList: true,
			subtree: true,
			attributes: true,
			characterData: true,
			attributeFilter: SELECT_OPTION_ATTRIBUTES,
		});
		this.choices = this.collect();
		super.connectedCallback();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.open = false;
		this.closePopup();
		const view = this.ownerDocument.defaultView;
		view?.clearTimeout(this.searchTimer);
		this.searchTimer = 0;
		this.search = '';
		this.observer?.disconnect();
		if (this.refreshFrame) view?.cancelAnimationFrame(this.refreshFrame);
		this.refreshFrame = 0;
	}

	private collect(): SelectChoice[] {
		const children = [...this.querySelectorAll('option')];
		if (children.length) return collectSelectChoices(children);
		return this.data.map((option) => ({
			value: option.value,
			label: option.label,
			disabled: option.disabled === true,
			hidden: false,
			group: '',
		}));
	}

	/** Re-read the choices after the `<option>` children or the `options` data changed. */
	private refresh(): void {
		this.choices = this.collect();
		this.requestUpdate();
	}

	private isUsable(choice: SelectChoice | undefined): boolean {
		return Boolean(choice) && !choice!.disabled && !choice!.hidden;
	}

	private show(): void {
		this.choices = this.collect();
		if (present(this.disabled) || this.open || !this.choices.some((c) => this.isUsable(c))) return;
		const selected = this.choices.findIndex((c) => c.value === this.value && this.isUsable(c));
		this.active = selected >= 0 ? selected : this.choices.findIndex((c) => this.isUsable(c));
		this.opened = true;
		this.open = true;
	}

	private close(): void {
		this.open = false;
	}

	/** The popup is a top-layer popover: show or hide it as `open` changes. */
	private openPopup(): void {
		const menu = this.menu;
		if (!menu) return;
		if (typeof menu.showPopover === 'function') menu.showPopover();
		this.position();
		const view = this.ownerDocument.defaultView;
		this.ownerDocument.addEventListener('pointerdown', this.onOutside, true);
		view?.addEventListener('resize', this.position);
		view?.addEventListener('scroll', this.position, true);
	}

	private closePopup(): void {
		const view = this.ownerDocument.defaultView;
		view?.clearTimeout(this.searchTimer);
		this.search = '';
		try {
			if (this.menu?.matches(':popover-open')) this.menu.hidePopover();
		} catch {
			/* No popover support. */
		}
		this.ownerDocument.removeEventListener('pointerdown', this.onOutside, true);
		view?.removeEventListener('resize', this.position);
		view?.removeEventListener('scroll', this.position, true);
	}

	private readonly onOutside = (event: Event): void => {
		if (!event.composedPath().includes(this)) this.close();
	};

	private readonly position = (): void => {
		const { menu, trigger } = this;
		if (this.open && menu && trigger) positionSelectMenu(menu, trigger, this);
	};

	private move(step: number): void {
		if (!this.open) this.show();
		if (this.open) this.active = nextSelectActive(this.choices, this.active, step);
	}

	private commit(index: number): void {
		const choice = this.choices[index];
		if (!choice || !this.isUsable(choice)) return;
		const changed = this.value !== choice.value;
		this.value = choice.value;
		this.close();
		this.trigger?.focus();
		if (changed) {
			this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
			this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
		}
	}

	private onKeyDown(event: KeyboardEvent): void {
		if (present(this.disabled)) return;
		const stop = () => {
			event.preventDefault();
			event.stopPropagation();
		};
		if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			stop();
			this.move(event.key === 'ArrowDown' ? 1 : -1);
		} else if (event.key === 'Home' || event.key === 'End') {
			stop();
			this.show();
			this.active =
				event.key === 'Home'
					? nextSelectActive(this.choices, -1, 1)
					: nextSelectActive(this.choices, 0, -1);
		} else if (event.key === 'PageUp' || event.key === 'PageDown') {
			stop();
			this.show();
			this.active = pageSelectActive(this.choices, this.active, event.key === 'PageDown' ? 1 : -1);
		} else if (event.key === 'Enter' || event.key === ' ') {
			stop();
			if (this.open) this.commit(this.active);
			else this.show();
		} else if (event.key === 'Escape' && this.open) {
			stop();
			this.close();
			this.trigger?.focus();
		} else if (event.key === 'Tab') this.close();
		else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
			event.stopPropagation();
			this.typeAhead(event.key);
		}
	}

	private typeAhead(key: string): void {
		const view = this.ownerDocument.defaultView;
		this.search += key.toLocaleLowerCase();
		view?.clearTimeout(this.searchTimer);
		this.searchTimer =
			view?.setTimeout(() => {
				this.search = '';
			}, 700) ?? 0;
		this.show();
		const index = this.choices.findIndex(
			(choice) => this.isUsable(choice) && choice.label.toLocaleLowerCase().startsWith(this.search),
		);
		if (index >= 0) this.active = index;
	}

	protected override willUpdate(): void {
		// Options may have changed since the last update (the observer only runs per frame).
		this.choices = this.collect();
		// A select without a value takes the default choice, like a native one.
		if (this.requested === null && this.choices.length) {
			this.requested = this.value;
			this.requestUpdate('value', undefined);
		}
		this.internals?.setFormValue?.(present(this.disabled) ? null : this.value);
		if (present(this.disabled) && this.open) this.open = false;
	}

	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('open')) {
			if (this.open) this.openPopup();
			else if (changed.get('open') !== undefined) this.closePopup();
		}
		if (this.open) {
			this.position();
			this.menu
				?.querySelector<HTMLElement>('[data-active]')
				?.scrollIntoView?.({ block: 'nearest' });
		}
	}

	private optionTemplates() {
		let previousGroup = '';
		return this.choices.map((choice, index) => {
			if (choice.hidden) return '';
			const heading =
				choice.group && choice.group !== previousGroup
					? html`<div class="group" role="presentation">${choice.group}</div>`
					: '';
			previousGroup = choice.group;
			const text = choice.displayLabel ?? choice.label;
			const note = choice.description
				? html`<span class="description">${choice.description}</span>`
				: '';
			return html`${heading}<div
					class="option"
					role="option"
					id="${this.menuId}-${index}"
					data-index=${index}
					?data-active=${this.open && index === this.active}
					aria-selected=${String(choice.value === this.value)}
					aria-disabled=${String(choice.disabled)}
					style=${ifDefined(choice.fontFamily ? `font-family:${choice.fontFamily}` : undefined)}
					>${text}${note}</div
				>`;
		});
	}

	protected override render() {
		const value = this.value;
		const shown = this.choices.find((choice) => choice.value === value);
		const name = this.getAttribute('aria-label');
		const open = present(this.open);
		const native = typeof HTMLElement.prototype.showPopover === 'function';
		return html`
			<button
				part="trigger"
				type="button"
				role="combobox"
				aria-haspopup="listbox"
				aria-expanded=${String(open)}
				aria-controls=${this.menuId}
				aria-label=${ifDefined(name ?? undefined)}
				aria-activedescendant=${ifDefined(
					open && this.active >= 0 ? `${this.menuId}-${this.active}` : undefined,
				)}
				?disabled=${present(this.disabled)}
				@click=${() => (this.open ? this.close() : this.show())}
				@keydown=${this.onKeyDown}
			>
				<slot name="icon"></slot>
				<span class="value" part="value">${shown?.displayLabel ?? shown?.label ?? value}</span>
				<span class="chevron" part="indicator" aria-hidden="true">
					<svg
						viewBox="0 0 24 24"
						fill="none"
						stroke="currentColor"
						stroke-width="2"
						stroke-linecap="round"
						stroke-linejoin="round"
					>
						<path d="m6 9 6 6 6-6"></path>
					</svg>
				</span>
			</button>
			<div
				class="menu"
				id=${this.menuId}
				role="listbox"
				popover="manual"
				part="popup"
				?data-fallback-open=${open && !native}
				@pointerdown=${(event: Event) => {
					if ((event.target as Element).closest('[data-index]')) event.preventDefault();
				}}
				@click=${(event: Event) => {
					const target = (event.target as Element).closest<HTMLElement>('[data-index]');
					if (target) this.commit(Number(target.dataset.index));
				}}
			>
				${this.querySelector('[slot="custom"]') ? html`<slot name="custom"></slot>` : ''}
				${this.opened ? this.optionTemplates() : ''}
			</div>
		`;
	}
}

export const defineSelect = definer('office-ui-select', () => OfficeUiSelect);
