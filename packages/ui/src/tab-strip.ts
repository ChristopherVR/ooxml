import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export interface OfficeTab {
	id: string;
	label: string;
	/** Tooltip; defaults to the label. */
	title?: string;
}
export type OfficeTabSelectEvent = CustomEvent<{ id: string }>;

const CSS = `
:host { display: flex; align-items: stretch; min-width: 0; min-height: 28px; gap: 2px; padding: 0 6px;
	background: var(--office-surface, #f3f4f6); color: var(--office-muted-foreground, #6b7280);
	border-top: 1px solid var(--office-border, #d1d5db); font-size: 12px;
	font-family: var(--office-font, system-ui, sans-serif); }
.step { flex: none; display: inline-grid; place-items: center; width: var(--office-target-size, 26px);
	min-height: var(--office-target-size, 26px);
	padding: 0; border: 0; background: transparent; color: inherit; cursor: pointer; }
.step:disabled { opacity: .4; cursor: default; }
.add { font-size: 18px; line-height: 1; }
.add[hidden] { display: none; }
.list { display: flex; align-items: stretch; min-width: 0; overflow-x: auto; scrollbar-width: thin; }
.tab { flex: none; max-width: 220px; min-height: var(--office-target-size, 26px); margin-top: -1px;
	padding: 2px 14px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis;
	border: 1px solid transparent; border-top: 0; border-radius: 0 0 var(--office-radius, 4px) var(--office-radius, 4px);
	background: transparent; color: inherit; font: inherit; cursor: pointer; }
.tab:hover { color: var(--office-foreground, #1f2937); }
.tab[aria-selected="true"] { border-color: var(--office-border, #d1d5db); background: var(--office-background, #fff);
	color: var(--office-accent, #2563eb); font-weight: 600; box-shadow: inset 0 2px 0 var(--office-accent, #2563eb); }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: -2px; }
svg { width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.5;
	stroke-linecap: round; stroke-linejoin: round; }
@media (forced-colors: active) {
	:host { border-top-color: CanvasText; }
	.tab[aria-selected="true"] { background: Highlight; color: HighlightText; forced-color-adjust: none; }
}
`;

/**
 * Bottom document tabs (Visio pages, Excel sheets). Set `tabs` (`OfficeTab[]`) and `selected`
 * (a tab id). Labels are inserted as text. The previous/next buttons and arrow, Home and End
 * keys move the selection; the tab list uses a roving tabindex. User selection emits
 * `office-tab-select` `{ id }` (cancelable); setting properties never emits.
 * Attributes: `label` (tab list name, default "Tabs"), `previous-label` and `next-label`
 * (step button names, default "Previous" and "Next"), `disabled`, and `add-label` to show an
 * add button after the tabs (Insert Page, New Sheet) that emits `office-command`
 * `{ command: 'tab-add' }`; `add-disabled` and `add-title` disable and explain it.
 */
export const defineTabStrip = definer('office-ui-tab-strip', () => {
	class OfficeUiTabStrip extends HTMLElement {
		static observedAttributes = [
			'label',
			'disabled',
			'selected',
			'previous-label',
			'next-label',
			'add-label',
			'add-disabled',
		];
		private items: OfficeTab[] = [];
		private readonly list: HTMLDivElement;
		private readonly previous: HTMLButtonElement;
		private readonly next: HTMLButtonElement;
		private readonly add: HTMLButtonElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const step = (name: string, icon: string) => {
				const button = doc.createElement('button');
				button.type = 'button';
				button.className = 'step';
				button.setAttribute('aria-label', name);
				button.title = name;
				const svg = createIconSvg(doc);
				paintIcon(svg, icon);
				button.append(svg);
				return button;
			};
			this.previous = step('Previous', 'chevronLeft');
			this.next = step('Next', 'chevronRight');
			this.list = doc.createElement('div');
			this.list.className = 'list';
			this.list.setAttribute('role', 'tablist');
			this.add = doc.createElement('button');
			this.add.type = 'button';
			this.add.className = 'step add';
			this.add.textContent = '+';
			this.add.addEventListener('click', () => {
				if (!this.add.disabled) emit(this, 'office-command', { command: 'tab-add' });
			});
			root.append(this.previous, this.next, this.list, this.add);
			this.previous.addEventListener('click', () => this.move(-1, false));
			this.next.addEventListener('click', () => this.move(1, false));
			this.list.addEventListener('click', (event) => {
				const tab = (event.target as Element).closest<HTMLButtonElement>('.tab');
				if (tab?.dataset.id !== undefined) this.choose(tab.dataset.id, false);
			});
			this.list.addEventListener('keydown', (event) => this.onKey(event));
		}
		connectedCallback(): void {
			this.render();
		}
		attributeChangedCallback(name: string): void {
			if (name === 'selected') this.sync();
			else this.render();
		}
		get tabs(): OfficeTab[] {
			return this.items.map((tab) => ({ ...tab }));
		}
		set tabs(value: readonly OfficeTab[]) {
			this.items = (Array.isArray(value) ? value : [])
				.filter((tab) => tab && typeof tab.id === 'string')
				.map((tab) => ({
					id: tab.id,
					label: String(tab.label ?? ''),
					...(tab.title === undefined ? {} : { title: String(tab.title) }),
				}));
			this.render();
		}
		get selected(): string {
			return this.getAttribute('selected') ?? '';
		}
		set selected(id: string) {
			this.setAttribute('selected', String(id));
		}
		private index(): number {
			return this.items.findIndex((tab) => tab.id === this.selected);
		}
		private move(delta: number, focus: boolean): void {
			const index = this.index();
			const target = this.items[Math.max(0, Math.min(this.items.length - 1, index + delta))];
			if (target) this.choose(target.id, focus);
		}
		private choose(id: string, focus: boolean): void {
			if (this.hasAttribute('disabled')) return;
			if (id !== this.selected && emit(this, 'office-tab-select', { id }, true)) this.selected = id;
			if (focus) this.tabButton(this.selected)?.focus();
		}
		private tabButton(id: string): HTMLButtonElement | undefined {
			return [...this.list.querySelectorAll<HTMLButtonElement>('.tab')].find(
				(tab) => tab.dataset.id === id,
			);
		}
		private onKey(event: KeyboardEvent): void {
			const keys: Record<string, number> = {
				ArrowLeft: -1,
				ArrowRight: 1,
				ArrowUp: -1,
				ArrowDown: 1,
			};
			if (event.key === 'Home' || event.key === 'End') {
				event.preventDefault();
				const target = event.key === 'Home' ? this.items[0] : this.items[this.items.length - 1];
				if (target) this.choose(target.id, true);
			} else if (keys[event.key] !== undefined) {
				event.preventDefault();
				this.move(keys[event.key]!, true);
			}
		}
		private render(): void {
			const doc = this.ownerDocument;
			this.list.setAttribute('aria-label', this.getAttribute('label') ?? 'Tabs');
			for (const [button, name, fallback] of [
				[this.previous, 'previous-label', 'Previous'],
				[this.next, 'next-label', 'Next'],
			] as const) {
				const label = this.getAttribute(name) || fallback;
				button.setAttribute('aria-label', label);
				button.title = label;
			}
			this.list.replaceChildren(
				...this.items.map((item) => {
					const tab = doc.createElement('button');
					tab.type = 'button';
					tab.className = 'tab';
					tab.setAttribute('role', 'tab');
					tab.dataset.id = item.id;
					tab.textContent = item.label;
					tab.title = item.title ?? item.label;
					return tab;
				}),
			);
			this.sync();
		}
		private sync(): void {
			const disabled = this.hasAttribute('disabled');
			const index = this.index();
			for (const tab of this.list.querySelectorAll<HTMLButtonElement>('.tab')) {
				const selected = tab.dataset.id === this.selected;
				tab.setAttribute('aria-selected', String(selected));
				tab.tabIndex = selected || (index < 0 && tab === this.list.firstElementChild) ? 0 : -1;
				tab.disabled = disabled;
				if (selected) tab.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
			}
			this.previous.disabled = disabled || index <= 0;
			const addLabel = this.getAttribute('add-label');
			this.add.hidden = !addLabel;
			this.add.disabled = disabled || this.hasAttribute('add-disabled');
			this.add.setAttribute('aria-label', addLabel ?? 'Add');
			this.add.title = this.getAttribute('add-title') ?? addLabel ?? '';
			this.next.disabled = disabled || index < 0 || index >= this.items.length - 1;
		}
	}
	return OfficeUiTabStrip;
});
