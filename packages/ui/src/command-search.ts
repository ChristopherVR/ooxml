import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/** One searchable command: label and optional keywords; disabled commands show their reason. */
export interface OfficeSearchCommand {
	id: string;
	label: string;
	keywords?: string;
	/** Where the command lives, shown under the label (for example "Home › Paragraph"). */
	description?: string;
	disabled?: boolean;
	title?: string;
}

const CSS = `
:host { display: inline-flex; align-items: center; min-width: 0; font-family: ${tok('--office-font')}; }
.field { display: flex; align-items: center; gap: ${tok('--office-space-1-5')}; width: 100%; min-width: 0; height: ${tok('--office-target-size')};
	padding: 0 ${tok('--office-space-2')}; box-sizing: border-box; border: ${tok('--office-border-width')} solid transparent; border-radius: ${tok('--office-radius')};
	color: ${tok('--office-muted-foreground')}; }
.field:hover, .field:focus-within { border-color: ${tok('--office-border')}; background: ${tok('--office-background')}; }
input { flex: 1; min-width: 0; border: 0; outline: 0; background: transparent; color: ${tok('--office-foreground')};
	font: inherit; font-size: ${tok('--office-font-size')}; }
svg { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; flex: none; fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')};
	stroke-linecap: round; stroke-linejoin: round; }
ul { position: fixed; inset: auto; margin: 0; padding: ${tok('--office-space-1')}; min-width: ${tok('--office-search-popup-min-width')}; max-height: ${tok('--office-search-popup-max-height')}; overflow: auto;
	box-sizing: border-box; list-style: none; background: ${tok('--office-background')}; color: ${tok('--office-foreground')};
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius')};
	box-shadow: ${tok('--office-shadow')}; font-size: ${tok('--office-font-size-sm')}; }
ul[hidden] { display: none; }
li { display: flex; flex-direction: column; justify-content: center; min-height: ${tok('--office-target-size')};
	padding: ${tok('--office-space-0')} ${tok('--office-space-2-5')}; border-radius: ${tok('--office-radius-sm')}; cursor: pointer; }
li[data-active] { background: ${tok('--office-surface')}; }
li[aria-disabled="true"] { cursor: not-allowed; color: ${tok('--office-muted-foreground')}; }
li small { color: ${tok('--office-muted-foreground')}; font-size: ${tok('--office-font-size-xs')}; }
li.empty { cursor: default; color: ${tok('--office-muted-foreground')}; }
@media (forced-colors: active) {
	.field, ul { border-color: CanvasText; background: Canvas; color: CanvasText; }
	li[data-active] { background: Highlight; color: HighlightText; }
}
`;

let uid = 0;

/**
 * Office's "Tell me what you want to do" box. Set `commands` (`OfficeSearchCommand[]`); typing
 * lists up to `limit` (default 8) matches by label or keywords in a top-layer listbox.
 * ArrowUp/Down move, Enter or click run an enabled match (`office-command` `{ command }`),
 * Escape closes then clears. Disabled matches stay visible with their reason. Attributes:
 * `placeholder`, `label` (accessible name), `limit`. Setting `commands` never emits.
 */
export const defineCommandSearch = definer('office-ui-command-search', () => {
	class OfficeUiCommandSearch extends HTMLElement {
		static observedAttributes = ['placeholder', 'label'];
		private items: OfficeSearchCommand[] = [];
		private found: OfficeSearchCommand[] = [];
		private active = -1;
		private readonly input: HTMLInputElement;
		private readonly list: HTMLUListElement;
		private readonly uid = `office-command-search-${++uid}`;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(CSS));
			const field = doc.createElement('div');
			field.className = 'field';
			const glass = createIconSvg(doc);
			paintIcon(glass, 'search');
			glass.setAttribute('data-painted', '');
			this.input = doc.createElement('input');
			this.input.type = 'search';
			this.input.autocomplete = 'off';
			this.input.spellcheck = false;
			this.input.setAttribute('role', 'combobox');
			this.input.setAttribute('aria-autocomplete', 'list');
			this.input.setAttribute('aria-expanded', 'false');
			this.input.setAttribute('aria-controls', `${this.uid}-list`);
			this.list = doc.createElement('ul');
			this.list.id = `${this.uid}-list`;
			this.list.setAttribute('role', 'listbox');
			if (typeof (this.list as { showPopover?: unknown }).showPopover === 'function')
				this.list.setAttribute('popover', 'manual');
			else this.list.hidden = true;
			field.append(glass, this.input);
			root.append(field, this.list);
			this.input.addEventListener('input', () => this.update());
			this.input.addEventListener('keydown', (event) => this.onKey(event));
			this.input.addEventListener('blur', () => this.hide());
			// Pointer down keeps focus in the field so the choice is not lost to blur.
			this.list.addEventListener('pointerdown', (event) => event.preventDefault());
			this.list.addEventListener('click', (event) => {
				const item = (event.target as Element).closest<HTMLElement>('li[data-index]');
				if (item) this.choose(Number(item.dataset.index));
			});
		}
		connectedCallback(): void {
			this.sync();
		}
		disconnectedCallback(): void {
			this.hide();
		}
		attributeChangedCallback(): void {
			this.sync();
		}
		get commands(): OfficeSearchCommand[] {
			return this.items.map((item) => ({ ...item }));
		}
		set commands(value: readonly OfficeSearchCommand[]) {
			this.items = (Array.isArray(value) ? value : [])
				.filter((item) => item && typeof item.id === 'string')
				.map((item) => ({ ...item, label: String(item.label ?? '') }));
			if (this.isOpen) this.update();
		}
		get isOpen(): boolean {
			return this.input.getAttribute('aria-expanded') === 'true';
		}
		private sync(): void {
			this.input.placeholder = this.getAttribute('placeholder') ?? 'Tell me what you want to do';
			this.input.setAttribute('aria-label', this.getAttribute('label') ?? this.input.placeholder);
		}
		private update(): void {
			const query = this.input.value.trim().toLowerCase();
			if (!query) return this.hide();
			const limit = Math.max(1, Number(this.getAttribute('limit')) || 8);
			// Office ranks usable commands first, then label prefix matches, then the rest.
			const rank = (item: OfficeSearchCommand) =>
				(item.disabled ? 2 : 0) + (item.label.toLowerCase().startsWith(query) ? 0 : 1);
			this.found = this.items
				.map((item, order) => ({ item, order }))
				.filter(({ item }) =>
					`${item.label} ${item.keywords ?? ''} ${item.description ?? ''}`
						.toLowerCase()
						.includes(query),
				)
				.sort((a, b) => rank(a.item) - rank(b.item) || a.order - b.order)
				.map(({ item }) => item)
				.slice(0, limit);
			this.active = this.found.findIndex((item) => !item.disabled);
			this.render();
			this.show();
		}
		private render(): void {
			const doc = this.ownerDocument;
			if (!this.found.length) {
				const empty = doc.createElement('li');
				empty.className = 'empty';
				empty.setAttribute('role', 'option');
				empty.setAttribute('aria-disabled', 'true');
				empty.textContent = 'No matching commands';
				this.list.replaceChildren(empty);
				this.input.removeAttribute('aria-activedescendant');
				return;
			}
			this.list.replaceChildren(
				...this.found.map((item, index) => {
					const option = doc.createElement('li');
					option.id = `${this.uid}-${index}`;
					option.dataset.index = String(index);
					option.setAttribute('role', 'option');
					option.setAttribute('aria-selected', String(index === this.active));
					if (index === this.active) option.setAttribute('data-active', '');
					const label = doc.createElement('span');
					label.textContent = item.label;
					option.append(label);
					if (item.disabled) option.setAttribute('aria-disabled', 'true');
					const detail = item.disabled ? (item.title ?? 'Not available') : item.description;
					if (detail) {
						const line = doc.createElement('small');
						line.textContent = detail;
						option.append(line);
					}
					return option;
				}),
			);
			if (this.active >= 0)
				this.input.setAttribute('aria-activedescendant', `${this.uid}-${this.active}`);
			else this.input.removeAttribute('aria-activedescendant');
		}
		private show(): void {
			const box = this.input.getBoundingClientRect();
			this.list.style.left = `${Math.round(box.left)}px`;
			this.list.style.top = `${Math.round(box.bottom + 4)}px`;
			if (this.list.hasAttribute('popover')) {
				if (!this.isOpen) this.list.showPopover?.();
			} else this.list.hidden = false;
			this.input.setAttribute('aria-expanded', 'true');
		}
		private hide(): void {
			if (!this.isOpen) return;
			if (this.list.hasAttribute('popover')) {
				try {
					this.list.hidePopover?.();
				} catch {
					/* Already hidden. */
				}
			} else this.list.hidden = true;
			this.input.setAttribute('aria-expanded', 'false');
			this.input.removeAttribute('aria-activedescendant');
		}
		private move(step: number): void {
			const enabled = this.found
				.map((item, index) => (item.disabled ? -1 : index))
				.filter((i) => i >= 0);
			if (!enabled.length) return;
			const at = enabled.indexOf(this.active);
			this.active = enabled[(at + step + enabled.length) % enabled.length]!;
			this.render();
		}
		private choose(index: number): void {
			const item = this.found[index];
			if (!item || item.disabled) return;
			this.input.value = '';
			this.hide();
			emit(this, 'office-command', { command: item.id });
		}
		private onKey(event: KeyboardEvent): void {
			if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
				event.preventDefault();
				if (!this.isOpen) this.update();
				else this.move(event.key === 'ArrowDown' ? 1 : -1);
			} else if (event.key === 'Enter') {
				event.preventDefault();
				this.choose(this.active);
			} else if (event.key === 'Escape') {
				event.preventDefault();
				event.stopPropagation();
				if (this.isOpen) this.hide();
				else this.input.value = '';
			}
		}
	}
	return OfficeUiCommandSearch;
});
