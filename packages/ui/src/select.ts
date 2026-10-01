import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export interface OfficeSelectOption {
	value: string;
	label: string;
	disabled?: boolean;
}

const CSS = `
:host { display: inline-block; position: relative; vertical-align: middle; }
button { box-sizing: border-box; display: flex; align-items: center; justify-content: space-between; gap: 8px;
	width: 100%; min-height: var(--office-target-size, 28px); padding: 0 8px; font: inherit; font-size: 12px;
	color: var(--office-foreground, #1f2937); background: var(--office-background, #fff);
	border: 1px solid var(--office-border, #d1d5db); border-radius: var(--office-radius, 4px); cursor: pointer; }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 1px; }
:host([disabled]) button { opacity: .5; cursor: not-allowed; }
ul { position: absolute; z-index: 10; inset-inline: 0; top: 100%; margin: 2px 0 0; padding: 2px; list-style: none;
	max-height: 240px; overflow: auto; background: var(--office-background, #fff);
	border: 1px solid var(--office-border, #d1d5db); border-radius: var(--office-radius, 4px); }
ul[hidden] { display: none; }
li { display: flex; align-items: center; min-height: var(--office-target-size, 28px); padding: 0 8px;
	font-size: 12px; color: var(--office-foreground, #1f2937); cursor: pointer; border-radius: 3px; }
li[aria-selected="true"] { font-weight: 600; }
li[data-active] { background: var(--office-surface, #f3f4f6); }
li[aria-disabled="true"] { opacity: .5; cursor: not-allowed; }
@media (forced-colors: active) {
	button, ul { border-color: CanvasText; background: Canvas; color: CanvasText; }
	li[data-active] { background: Highlight; color: HighlightText; }
}
`;

let uid = 0;

/**
 * Single-select dropdown. Properties: `options` (array), `value`, `disabled`, `selectedIndex`;
 * attributes `value`, `disabled`, `aria-label`. Trigger has `role="combobox"` and the popup is a
 * `listbox`. Keyboard: ArrowDown/Up (open, move), Home, End, Enter/Space (choose), Escape
 * (close), printable keys (typeahead); disabled options are skipped. `input` then `change`
 * bubble when the user picks; setting a property is silent.
 */
export const defineSelect = definer('office-ui-select', () => {
	class OfficeUiSelect extends HTMLElement {
		static observedAttributes = ['value', 'disabled', 'aria-label'];
		private opts: OfficeSelectOption[] = [];
		private current = '';
		private active = -1;
		private typed = '';
		private typedAt = 0;
		private readonly trigger: HTMLButtonElement;
		private readonly list: HTMLUListElement;
		private readonly label: HTMLSpanElement;
		private readonly domId = `office-ui-select-${++uid}`;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(CSS));
			this.trigger = doc.createElement('button');
			this.trigger.type = 'button';
			this.trigger.setAttribute('role', 'combobox');
			this.trigger.setAttribute('aria-haspopup', 'listbox');
			this.trigger.setAttribute('aria-expanded', 'false');
			this.trigger.setAttribute('aria-controls', this.domId);
			this.label = doc.createElement('span');
			const caret = doc.createElement('span');
			caret.setAttribute('aria-hidden', 'true');
			caret.textContent = '▾';
			this.trigger.append(this.label, caret);
			this.list = doc.createElement('ul');
			this.list.id = this.domId;
			this.list.setAttribute('role', 'listbox');
			this.list.hidden = true;
			root.append(this.trigger, this.list);
			this.trigger.addEventListener('click', () => (this.isOpen ? this.close() : this.open()));
			this.trigger.addEventListener('keydown', (event) => this.onKey(event));
			this.trigger.addEventListener('blur', () => this.close());
			this.list.addEventListener('mousedown', (event) => event.preventDefault());
			this.list.addEventListener('click', (event) => {
				const item = (event.target as Element).closest('li');
				if (item) this.choose(Number(item.dataset.index));
			});
		}
		connectedCallback(): void {
			this.sync();
		}
		attributeChangedCallback(name: string, _old: string | null, value: string | null): void {
			if (name === 'value') this.current = value ?? '';
			this.sync();
		}
		get options(): OfficeSelectOption[] {
			return this.opts;
		}
		set options(next: OfficeSelectOption[]) {
			this.opts = next.map((option) => ({ ...option }));
			this.sync();
		}
		get value(): string {
			return this.current;
		}
		set value(next: string) {
			this.current = String(next);
			this.sync();
		}
		get selectedIndex(): number {
			return this.opts.findIndex((option) => option.value === this.current);
		}
		set selectedIndex(index: number) {
			const option = this.opts[index];
			if (option) this.value = option.value;
		}
		get disabled(): boolean {
			return this.hasAttribute('disabled');
		}
		set disabled(value: boolean) {
			this.toggleAttribute('disabled', Boolean(value));
		}
		private get isOpen(): boolean {
			return !this.list.hidden;
		}
		private open(): void {
			if (this.disabled || this.isOpen) return;
			this.list.hidden = false;
			this.trigger.setAttribute('aria-expanded', 'true');
			this.setActive(this.selectedIndex >= 0 ? this.selectedIndex : this.step(-1, 1));
		}
		private close(): void {
			this.list.hidden = true;
			this.trigger.setAttribute('aria-expanded', 'false');
			this.trigger.removeAttribute('aria-activedescendant');
			this.active = -1;
		}
		private step(from: number, dir: 1 | -1): number {
			for (let i = from + dir; i >= 0 && i < this.opts.length; i += dir) {
				if (!this.opts[i]?.disabled) return i;
			}
			return from;
		}
		private setActive(index: number): void {
			this.active = index;
			this.list
				.querySelectorAll('li')
				.forEach((item, i) => item.toggleAttribute('data-active', i === index));
			if (index >= 0) this.trigger.setAttribute('aria-activedescendant', `${this.domId}-${index}`);
		}
		private choose(index: number): void {
			const option = this.opts[index];
			if (!option || option.disabled) return;
			const changed = option.value !== this.current;
			this.value = option.value;
			this.close();
			if (changed) {
				this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
				this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
			}
		}
		private onKey(event: KeyboardEvent): void {
			if (this.disabled) return;
			const key = event.key;
			if (key === 'ArrowDown' || key === 'ArrowUp') {
				event.preventDefault();
				if (!this.isOpen) return this.open();
				this.setActive(this.step(this.active, key === 'ArrowDown' ? 1 : -1));
			} else if (key === 'Home' || key === 'End') {
				event.preventDefault();
				if (!this.isOpen) this.open();
				this.setActive(key === 'Home' ? this.step(-1, 1) : this.step(this.opts.length, -1));
			} else if (key === 'Enter' || key === ' ') {
				event.preventDefault();
				if (this.isOpen) this.choose(this.active);
				else this.open();
			} else if (key === 'Escape') {
				if (this.isOpen) {
					event.preventDefault();
					event.stopPropagation();
					this.close();
				}
			} else if (key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
				const now = Date.now();
				this.typed = now - this.typedAt > 700 ? key : this.typed + key;
				this.typedAt = now;
				const needle = this.typed.toLowerCase();
				const index = this.opts.findIndex(
					(option) => !option.disabled && option.label.toLowerCase().startsWith(needle),
				);
				if (index >= 0) {
					if (this.isOpen) this.setActive(index);
					else this.choose(index);
				}
			}
		}
		private sync(): void {
			this.trigger.disabled = this.disabled;
			this.setAttribute('aria-disabled', String(this.disabled));
			const name = this.getAttribute('aria-label');
			if (name) this.trigger.setAttribute('aria-label', name);
			else this.trigger.removeAttribute('aria-label');
			this.label.textContent = this.opts[this.selectedIndex]?.label ?? '';
			this.list.replaceChildren(
				...this.opts.map((option, index) => {
					const item = this.ownerDocument.createElement('li');
					item.id = `${this.domId}-${index}`;
					item.dataset.index = String(index);
					item.setAttribute('role', 'option');
					item.setAttribute('aria-selected', String(option.value === this.current));
					if (option.disabled) item.setAttribute('aria-disabled', 'true');
					item.textContent = option.label;
					return item;
				}),
			);
			if (this.isOpen) this.setActive(Math.min(this.active, this.opts.length - 1));
		}
	}
	return OfficeUiSelect;
});
