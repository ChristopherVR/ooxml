import { definer } from './registry.js';
import {
	collectSelectChoices,
	createSelectTrigger,
	markSelectActive,
	nextSelectActive,
	pageSelectActive,
	positionSelectMenu,
	prependSelectCustomSlot,
	renderSelectMenu,
	SELECT_OPTION_ATTRIBUTES,
	type SelectChoice,
} from './select-menu.js';
import { SELECT_CSS } from './select-styles.js';
import { attachStyles, controlCss } from './styles.js';

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
export const defineSelect = definer('office-ui-select', () => {
	class OfficeUiSelect extends HTMLElement {
		static formAssociated = true;
		static observedAttributes = ['value', 'disabled', 'aria-label'];
		private readonly internals: ElementInternals | undefined;
		private readonly trigger: HTMLButtonElement;
		private readonly text: HTMLSpanElement;
		private readonly menu: HTMLDivElement;
		private readonly observer: MutationObserver | undefined;
		private refreshFrame = 0;
		private data: OfficeSelectOption[] = [];
		private choices: SelectChoice[] = [];
		private active = -1;
		private isOpen = false;
		private search = '';
		private searchTimer = 0;
		private defaultValue = '';
		private menuKey = '';

		constructor() {
			super();
			try {
				this.internals = this.attachInternals();
			} catch {
				this.internals = undefined;
			}
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(SELECT_CSS));
			({ trigger: this.trigger, text: this.text } = createSelectTrigger(doc));
			this.menu = doc.createElement('div');
			this.menu.className = 'menu';
			this.menu.id = `office-ui-select-options-${++uid}`;
			this.menu.setAttribute('role', 'listbox');
			this.menu.setAttribute('popover', 'manual');
			this.menu.setAttribute('part', 'popup');
			this.trigger.setAttribute('aria-controls', this.menu.id);
			root.append(this.trigger, this.menu);
			this.trigger.addEventListener('click', () => (this.isOpen ? this.close() : this.show()));
			this.trigger.addEventListener('keydown', (event) => this.onKeyDown(event));
			this.menu.addEventListener('pointerdown', (event) => {
				if ((event.target as Element).closest('[data-index]')) event.preventDefault();
			});
			this.menu.addEventListener('click', (event) => {
				const target = (event.target as Element).closest<HTMLElement>('[data-index]');
				if (target) this.commit(Number(target.dataset.index));
			});
			const Observer = doc.defaultView?.MutationObserver;
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

		connectedCallback(): void {
			this.defaultValue = this.getAttribute('value') ?? '';
			this.observer?.observe(this, {
				childList: true,
				subtree: true,
				attributes: true,
				characterData: true,
				attributeFilter: SELECT_OPTION_ATTRIBUTES,
			});
			this.refresh();
		}
		disconnectedCallback(): void {
			this.close();
			const view = this.ownerDocument.defaultView;
			view?.clearTimeout(this.searchTimer);
			this.searchTimer = 0;
			this.search = '';
			this.observer?.disconnect();
			if (this.refreshFrame) view?.cancelAnimationFrame(this.refreshFrame);
			this.refreshFrame = 0;
		}
		attributeChangedCallback(): void {
			this.refresh();
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
			const requested = this.getAttribute('value');
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
			this.setAttribute('value', String(next ?? ''));
		}
		get selectedIndex(): number {
			return this.options.findIndex((option) => option.value === this.value);
		}
		set selectedIndex(index: number) {
			this.value = this.options[index]?.value ?? '';
		}
		get disabled(): boolean {
			return this.hasAttribute('disabled');
		}
		set disabled(next: boolean) {
			this.toggleAttribute('disabled', Boolean(next));
		}
		override focus(options?: FocusOptions): void {
			this.trigger.focus(options);
		}
		formResetCallback(): void {
			this.value = this.defaultValue;
		}
		formDisabledCallback(disabled: boolean): void {
			this.disabled = disabled;
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

		private refresh(): void {
			if (!this.trigger) return;
			this.choices = this.collect();
			if (!this.hasAttribute('value') && this.choices.length) {
				const children = [...this.querySelectorAll('option')];
				this.setAttribute(
					'value',
					children.find((option) => option.selected)?.value ??
						this.choices.find((choice) => !choice.disabled && !choice.hidden)?.value ??
						this.choices[0]!.value,
				);
				return; // attributeChangedCallback refreshes again with the value set.
			}
			const shown = this.choices.find((choice) => choice.value === this.value);
			this.text.textContent = shown?.displayLabel ?? shown?.label ?? this.value;
			this.trigger.disabled = this.disabled;
			if (this.disabled) this.close();
			const name = this.getAttribute('aria-label');
			if (name) this.trigger.setAttribute('aria-label', name);
			else this.trigger.removeAttribute('aria-label');
			this.internals?.setFormValue?.(this.disabled ? null : this.value);
			// Hosts re-sync the same values often; rebuilding an open popup under the pointer would
			// swallow the click on an option.
			if ((this.isOpen || this.menu.childElementCount) && this.menuKey !== this.currentMenuKey())
				this.renderMenu();
		}

		private currentMenuKey(): string {
			return `${this.value}|${JSON.stringify(this.choices)}`;
		}

		private renderMenu(): void {
			this.menuKey = this.currentMenuKey();
			renderSelectMenu(this.menu, this.choices, this.value);
			prependSelectCustomSlot(this, this.menu);
			markSelectActive(this.menu, this.trigger, this.active, this.isOpen);
		}

		private show(): void {
			if (this.disabled || this.isOpen || !this.choices.some((c) => !c.disabled && !c.hidden))
				return;
			this.isOpen = true;
			this.setAttribute('open', '');
			const selected = this.choices.findIndex(
				(choice) => choice.value === this.value && !choice.disabled && !choice.hidden,
			);
			this.active =
				selected >= 0
					? selected
					: this.choices.findIndex((choice) => !choice.disabled && !choice.hidden);
			this.renderMenu();
			if (typeof this.menu.showPopover === 'function') this.menu.showPopover();
			else this.menu.dataset.fallbackOpen = '';
			this.trigger.setAttribute('aria-expanded', 'true');
			this.position();
			markSelectActive(this.menu, this.trigger, this.active, this.isOpen);
			const view = this.ownerDocument.defaultView;
			this.ownerDocument.addEventListener('pointerdown', this.onOutside, true);
			view?.addEventListener('resize', this.position);
			view?.addEventListener('scroll', this.position, true);
		}

		private close(): void {
			if (!this.isOpen) return;
			this.isOpen = false;
			const view = this.ownerDocument.defaultView;
			view?.clearTimeout(this.searchTimer);
			this.search = '';
			this.removeAttribute('open');
			try {
				if (this.menu.matches(':popover-open')) this.menu.hidePopover();
			} catch {
				/* No popover support. */
			}
			delete this.menu.dataset.fallbackOpen;
			this.trigger.setAttribute('aria-expanded', 'false');
			this.trigger.removeAttribute('aria-activedescendant');
			this.ownerDocument.removeEventListener('pointerdown', this.onOutside, true);
			view?.removeEventListener('resize', this.position);
			view?.removeEventListener('scroll', this.position, true);
		}

		private readonly onOutside = (event: Event): void => {
			if (!event.composedPath().includes(this)) this.close();
		};

		private readonly position = (): void => {
			if (this.isOpen) positionSelectMenu(this.menu, this.trigger, this);
		};

		private move(step: number): void {
			if (!this.isOpen) this.show();
			if (!this.isOpen) return;
			this.active = nextSelectActive(this.choices, this.active, step);
			markSelectActive(this.menu, this.trigger, this.active, this.isOpen);
		}

		private commit(index: number): void {
			const choice = this.choices[index];
			if (!choice || choice.disabled || choice.hidden) return;
			const changed = this.value !== choice.value;
			this.value = choice.value;
			this.close();
			this.trigger.focus();
			if (changed) {
				this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
				this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
			}
		}

		private onKeyDown(event: KeyboardEvent): void {
			if (this.disabled) return;
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
				markSelectActive(this.menu, this.trigger, this.active, this.isOpen);
			} else if (event.key === 'PageUp' || event.key === 'PageDown') {
				stop();
				this.show();
				this.active = pageSelectActive(
					this.choices,
					this.active,
					event.key === 'PageDown' ? 1 : -1,
				);
				markSelectActive(this.menu, this.trigger, this.active, this.isOpen);
			} else if (event.key === 'Enter' || event.key === ' ') {
				stop();
				if (this.isOpen) this.commit(this.active);
				else this.show();
			} else if (event.key === 'Escape' && this.isOpen) {
				stop();
				this.close();
				this.trigger.focus();
			} else if (event.key === 'Tab') this.close();
			else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
				event.stopPropagation();
				const view = this.ownerDocument.defaultView;
				this.search += event.key.toLocaleLowerCase();
				view?.clearTimeout(this.searchTimer);
				this.searchTimer =
					view?.setTimeout(() => {
						this.search = '';
					}, 700) ?? 0;
				this.show();
				const index = this.choices.findIndex(
					(choice) =>
						!choice.disabled &&
						!choice.hidden &&
						choice.label.toLocaleLowerCase().startsWith(this.search),
				);
				if (index >= 0) {
					this.active = index;
					markSelectActive(this.menu, this.trigger, this.active, this.isOpen);
				}
			}
		}
	}
	return OfficeUiSelect;
});
