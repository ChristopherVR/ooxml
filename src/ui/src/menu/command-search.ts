import { LitElement, html, type PropertyValues } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeElement, controlStyles } from '../base';
import { glyph } from '../glyph';
import { definer } from '../registry';
import css from './command-search.css?raw';

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

let uid = 0;

/**
 * Office's "Tell me what you want to do" box. Set `commands` (`OfficeSearchCommand[]`); typing
 * lists up to `limit` (default 8) matches by label or keywords in a top-layer listbox.
 * ArrowUp/Down move, Enter or click run an enabled match (`office-command` `{ command }`),
 * Escape closes then clears. Disabled matches stay visible with their reason. Attributes:
 * `placeholder`, `label` (accessible name), `limit`. Setting `commands` never emits.
 */
export class OfficeUiCommandSearch extends OfficeElement {
	static override shadowRootOptions = { ...LitElement.shadowRootOptions, delegatesFocus: true };
	static override styles = controlStyles(css);
	static override properties = {
		placeholder: { type: String },
		label: { type: String },
		limit: { type: String },
		found: { state: true },
		active: { state: true },
		isOpen: { state: true },
	};
	declare placeholder: string | null;
	declare label: string | null;
	declare limit: string | null;
	declare found: OfficeSearchCommand[];
	declare active: number;
	declare isOpen: boolean;
	private items: OfficeSearchCommand[] = [];
	private readonly uid = `office-command-search-${++uid}`;

	constructor() {
		super();
		this.placeholder = null;
		this.label = null;
		this.limit = null;
		this.found = [];
		this.active = -1;
		this.isOpen = false;
	}

	get commands(): OfficeSearchCommand[] {
		return this.items.map((item) => ({ ...item }));
	}
	set commands(value: readonly OfficeSearchCommand[]) {
		this.items = (Array.isArray(value) ? value : [])
			.filter((item) => item && typeof item.id === 'string')
			.map((item) => ({ ...item, label: String(item.label ?? '') }));
		if (this.isOpen) this.search();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.hide();
	}

	private get input(): HTMLInputElement | null {
		return this.renderRoot.querySelector('input');
	}
	private get list(): HTMLElement | null {
		return this.renderRoot.querySelector('ul');
	}

	private search(): void {
		const query = (this.input?.value ?? '').trim().toLowerCase();
		if (!query) return this.hide();
		const limit = Math.max(1, Number(this.limit) || 8);
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
		this.isOpen = true;
	}

	private hide(): void {
		this.isOpen = false;
	}

	private move(step: number): void {
		const enabled = this.found
			.map((item, index) => (item.disabled ? -1 : index))
			.filter((i) => i >= 0);
		if (!enabled.length) return;
		const at = enabled.indexOf(this.active);
		this.active = enabled[(at + step + enabled.length) % enabled.length]!;
	}

	private choose(index: number): void {
		const item = this.found[index];
		if (!item || item.disabled) return;
		if (this.input) this.input.value = '';
		this.hide();
		this.fire('office-command', { command: item.id });
	}

	private onKey(event: KeyboardEvent): void {
		if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			event.preventDefault();
			if (!this.isOpen) this.search();
			else this.move(event.key === 'ArrowDown' ? 1 : -1);
		} else if (event.key === 'Enter') {
			event.preventDefault();
			this.choose(this.active);
		} else if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			if (this.isOpen) this.hide();
			else if (this.input) this.input.value = '';
		}
	}

	/** The listbox is a top-layer popover below the field: show or hide it as `isOpen` changes. */
	protected override updated(changed: PropertyValues<this>): void {
		const list = this.list;
		if (!list || !changed.has('isOpen')) return;
		if (this.isOpen) {
			const box = this.input?.getBoundingClientRect();
			if (box) {
				list.style.left = `${Math.round(box.left)}px`;
				list.style.top = `${Math.round(box.bottom + 4)}px`;
			}
			if (list.hasAttribute('popover') && !list.matches(':popover-open')) list.showPopover?.();
		} else if (list.hasAttribute('popover')) {
			try {
				list.hidePopover?.();
			} catch {
				/* Already hidden. */
			}
		}
	}

	private optionTemplates() {
		if (!this.found.length)
			return html`<li class="empty" role="option" aria-disabled="true">No matching commands</li>`;
		return this.found.map((item, index) => {
			const detail = item.disabled ? (item.title ?? 'Not available') : item.description;
			return html`<li
				id="${this.uid}-${index}"
				data-index=${index}
				role="option"
				aria-selected=${String(index === this.active)}
				aria-disabled=${ifDefined(item.disabled ? 'true' : undefined)}
				?data-active=${index === this.active}
				><span>${item.label}</span>${detail ? html`<small>${detail}</small>` : ''}</li
			>`;
		});
	}

	protected override render() {
		const placeholder = this.placeholder ?? 'Tell me what you want to do';
		const popover = typeof HTMLElement.prototype.showPopover === 'function';
		const activeId =
			this.isOpen && this.found.length && this.active >= 0
				? `${this.uid}-${this.active}`
				: undefined;
		return html`
			<div class="field">
				${glyph('search', 'icon')}
				<input
					type="search"
					autocomplete="off"
					spellcheck="false"
					role="combobox"
					aria-autocomplete="list"
					aria-expanded=${String(this.isOpen)}
					aria-controls="${this.uid}-list"
					aria-label=${this.label ?? placeholder}
					aria-activedescendant=${ifDefined(activeId)}
					placeholder=${placeholder}
					@input=${this.search}
					@keydown=${this.onKey}
					@blur=${this.hide}
				/>
			</div>
			<ul
				id="${this.uid}-list"
				role="listbox"
				popover=${ifDefined(popover ? 'manual' : undefined)}
				?hidden=${!popover && !this.isOpen}
				@pointerdown=${(event: Event) => event.preventDefault()}
				@click=${(event: Event) => {
					const item = (event.target as Element).closest<HTMLElement>('li[data-index]');
					if (item) this.choose(Number(item.dataset.index));
				}}
				>${this.isOpen ? this.optionTemplates() : ''}</ul
			>
		`;
	}
}

export const defineCommandSearch = definer('office-ui-command-search', () => OfficeUiCommandSearch);
