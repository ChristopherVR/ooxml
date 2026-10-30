import { html, type PropertyValues } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { definer, present } from '../registry';
import css from './tab-strip.css?raw';
import { indexAfterDrop, wireTabDrag } from './tab-drag';

export interface OfficeTab {
	id: string;
	label: string;
	/** Tooltip; defaults to the label. */
	title?: string;
	/** Tabs only reorder within their own group (Visio keeps background pages after the others). */
	group?: string;
}
export type OfficeTabReorderEvent = CustomEvent<{ id: string; index: number }>;
export type OfficeTabSelectEvent = CustomEvent<{ id: string }>;

const KEY_STEPS: Record<string, number> = {
	ArrowLeft: -1,
	ArrowRight: 1,
	ArrowUp: -1,
	ArrowDown: 1,
};

/**
 * Bottom document tabs (Visio pages, Excel sheets). Set `tabs` (`OfficeTab[]`) and `selected`
 * (a tab id). Labels are inserted as text. The previous/next buttons and arrow, Home and End
 * keys move the selection; the tab list uses a roving tabindex. User selection emits
 * `office-tab-select` `{ id }` (cancelable); setting properties never emits.
 * Attributes: `label` (tab list name, default "Tabs"), `previous-label` and `next-label`
 * (step button names, default "Previous" and "Next"), `disabled`, and `add-label` to show an
 * add button after the tabs (Insert Page, New Sheet) that emits `office-command`
 * `{ command: 'tab-add' }`; `add-disabled` and `add-title` disable and explain it. The
 * `after-tabs` slot sits between the tabs and the add button (Visio's All pages list), and the
 * `reorderable` attribute lets a pointer (mouse, pen or touch) drag a tab along the strip: past a
 * small threshold an insertion marker shows the gap, Escape cancels, and releasing emits
 * `office-tab-reorder` `{ id, index }` (the tab's new index) without moving anything itself; the
 * host reorders its model and sets `tabs` again. A tab with a `group` only drops within that group.
 * The `plain` attribute draws the selected tab as bold accent text without a box, as Visio does.
 */
export class OfficeUiTabStrip extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		// `tabs` is sanitised and copied, `selected` reads as '' when unset: both have accessors.
		selected: { type: String, noAccessor: true },
		label: { type: String },
		previousLabel: { attribute: 'previous-label', type: String },
		nextLabel: { attribute: 'next-label', type: String },
		addLabel: { attribute: 'add-label', type: String },
		addTitle: { attribute: 'add-title', type: String },
		addDisabled: { attribute: 'add-disabled', ...flag },
		disabled: flag,
		reorderable: flag,
	};
	declare label: string | null;
	declare previousLabel: string | null;
	declare nextLabel: string | null;
	declare addLabel: string | null;
	declare addTitle: string | null;
	declare addDisabled: boolean;
	declare disabled: boolean;
	declare reorderable: boolean;
	private disposeDrag: (() => void) | undefined;
	private items: OfficeTab[] = [];
	private picked: string | null = null;

	constructor() {
		super();
		this.label = null;
		this.previousLabel = null;
		this.nextLabel = null;
		this.addLabel = null;
		this.addTitle = null;
		this.addDisabled = false;
		this.disabled = false;
		this.reorderable = false;
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
				...(tab.group === undefined ? {} : { group: String(tab.group) }),
			}));
		this.requestUpdate('tabs');
	}

	get selected(): string {
		return this.picked ?? '';
	}
	set selected(id: string | null) {
		this.picked = id === null ? null : String(id);
		this.reflectAttribute('selected', this.picked, {});
		this.requestUpdate('selected');
	}

	private index(): number {
		return this.items.findIndex((tab) => tab.id === this.selected);
	}

	private move(delta: number, focus: boolean): void {
		const target = this.items[Math.max(0, Math.min(this.items.length - 1, this.index() + delta))];
		if (target) this.choose(target.id, focus);
	}

	private choose(id: string, focus: boolean): void {
		if (present(this.disabled)) return;
		if (id !== this.selected && this.fire('office-tab-select', { id }, true)) this.selected = id;
		if (focus)
			[...this.renderRoot.querySelectorAll<HTMLButtonElement>('.tab')]
				.find((tab) => tab.dataset.id === this.selected)
				?.focus();
	}

	private onListClick(event: Event): void {
		const tab = (event.target as Element).closest<HTMLButtonElement>('.tab');
		if (tab?.dataset.id !== undefined) this.choose(tab.dataset.id, false);
	}

	private onKey(event: KeyboardEvent): void {
		if (event.key === 'Home' || event.key === 'End') {
			event.preventDefault();
			const target = event.key === 'Home' ? this.items[0] : this.items[this.items.length - 1];
			if (target) this.choose(target.id, true);
		} else if (KEY_STEPS[event.key] !== undefined) {
			event.preventDefault();
			this.move(KEY_STEPS[event.key]!, true);
		}
	}

	override connectedCallback(): void {
		super.connectedCallback();
		if (this.hasUpdated) this.wireDrag();
	}

	override disconnectedCallback(): void {
		this.disposeDrag?.();
		this.disposeDrag = undefined;
		super.disconnectedCallback();
	}

	private wireDrag(): void {
		const list = this.renderRoot.querySelector<HTMLElement>('.list');
		if (!list || this.disposeDrag) return;
		const tabs = () => [...list.querySelectorAll<HTMLElement>('.tab')];
		const groupOf = (tab: HTMLElement) =>
			this.items.find((item) => item.id === tab.dataset.id)?.group;
		this.disposeDrag = wireTabDrag({
			area: list,
			tabs,
			tabSelector: '.tab',
			canDrag: (tab) => present(this.reorderable) && !present(this.disabled) && !!tab.dataset.id,
			draggingClass: 'dragging',
			markerClass: 'marker',
			markerParent: this.renderRoot,
			markerBox: this,
			limit: (tab) => {
				const group = groupOf(tab);
				const same = tabs().filter((other) => groupOf(other) === group);
				const first = tabs().indexOf(same[0]!);
				return [first, first + same.length];
			},
			drop: (tab, slot) => {
				const from = tabs().indexOf(tab);
				const index = indexAfterDrop(from, slot);
				if (from >= 0 && index !== from)
					this.fire('office-tab-reorder', { id: tab.dataset.id!, index });
			},
		});
	}

	protected override firstUpdated(_changed: PropertyValues<this>): void {
		this.wireDrag();
	}

	protected override updated(_changed: PropertyValues<this>): void {
		this.renderRoot
			.querySelector<HTMLElement>('.tab[aria-selected="true"]')
			?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
	}

	protected override render() {
		const disabled = present(this.disabled);
		const index = this.index();
		const previous = this.previousLabel || 'Previous';
		const next = this.nextLabel || 'Next';
		return html`
			<button
				type="button"
				class="step"
				aria-label=${previous}
				title=${previous}
				?disabled=${disabled || index <= 0}
				@click=${() => this.move(-1, false)}
				>${glyph('chevronLeft', 'icon')}</button
			>
			<button
				type="button"
				class="step"
				aria-label=${next}
				title=${next}
				?disabled=${disabled || index < 0 || index >= this.items.length - 1}
				@click=${() => this.move(1, false)}
				>${glyph('chevronRight', 'icon')}</button
			>
			<div
				class="list"
				role="tablist"
				aria-label=${this.label ?? 'Tabs'}
				@click=${this.onListClick}
				@keydown=${this.onKey}
			>
				${this.items.map((item, at) => {
					const selected = item.id === this.selected;
					return html`<button
						type="button"
						class="tab"
						role="tab"
						data-id=${item.id}
						title=${item.title ?? item.label}
						aria-selected=${String(selected)}
						tabindex=${selected || (index < 0 && at === 0) ? 0 : -1}
						?disabled=${disabled}
						>${item.label}</button
					>`;
				})}
			</div>
			<slot name="after-tabs"></slot>
			<button
				type="button"
				class="step add"
				aria-label=${this.addLabel ?? 'Add'}
				title=${this.addTitle ?? this.addLabel ?? ''}
				?hidden=${!this.addLabel}
				?disabled=${disabled || present(this.addDisabled)}
				@click=${() => !present(this.addDisabled) && this.fire('office-command', { command: 'tab-add' })}
				>+</button
			>
		`;
	}
}

export const defineTabStrip = definer('office-ui-tab-strip', () => OfficeUiTabStrip);
