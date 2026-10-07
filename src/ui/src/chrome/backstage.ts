import { html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, present } from '../registry';
import css from './backstage.css?raw';

/**
 * Office's File view (backstage): a navigation column and one page per item. Generalised from
 * pptx-viewer's `BACKSTAGE_NAV` model (render/backstage.ts) and the Visio viewer's backstage.
 */
export interface OfficeBackstageItem {
	id: string;
	label: string;
	/** Items in the footer group sit at the bottom of the column (Account, Options). */
	group?: 'footer';
	disabled?: boolean;
	/** Tooltip, for example why an item is disabled. */
	title?: string;
}

/** An item was chosen; cancelable (`preventDefault()` keeps the current page). */
export type OfficeBackstageSelectEvent = CustomEvent<{ id: string }>;
export type OfficeBackstageCloseEvent = CustomEvent<{ reason: 'back' | 'escape' | 'api' }>;

/**
 * `<office-ui-backstage>`: set `items`, put one light-DOM child per page with
 * `data-backstage-page="<id>"`, then `show(id)`. Choosing an item emits `office-backstage-select`
 * and, unless prevented, shows its page; items without a page (Save, Close) only emit. Back and
 * Escape emit a cancelable `office-backstage-close` and close. Attributes: `open`, `selected`,
 * `label` (default "File"), `back-label`. Nav buttons carry `data-backstage-item`.
 */
export class OfficeUiBackstage extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		open: flag,
		items: { attribute: false },
		label: { type: String },
		backLabel: { attribute: 'back-label', type: String },
		// The pages hide against `selected`, which reads as '' when unset, so it has an accessor.
		selected: { type: String, noAccessor: true },
	};
	declare open: boolean;
	declare items: readonly OfficeBackstageItem[];
	declare label: string | null;
	declare backLabel: string | null;
	private picked: string | null = null;

	constructor() {
		super();
		this.open = false;
		this.items = [];
		this.label = null;
		this.backLabel = null;
		this.addEventListener('keydown', (event) => {
			if (!present(this.open)) return;
			if (event.key === 'Tab') return this.trapTab(event);
			if (event.key !== 'Escape') return;
			event.preventDefault();
			event.stopPropagation();
			this.requestClose('escape');
		});
	}

	/** The controls Tab visits while open, in order: back, the navigation, then the shown page. */
	private tabStops(): HTMLElement[] {
		const chrome = [...this.renderRoot.querySelectorAll<HTMLElement>('button:not([disabled])')];
		const page = [...this.children]
			.filter((child): child is HTMLElement => child instanceof HTMLElement && !child.hidden)
			.flatMap((child) => [
				...child.querySelectorAll<HTMLElement>(
					'a[href], button, input, select, textarea, [tabindex]:not([tabindex="-1"])',
				),
			])
			.filter((control) => !(control as HTMLButtonElement).disabled);
		return [...chrome, ...page].filter((control) => control.getClientRects().length > 0);
	}

	/** Keeps Tab inside the open File view: past the last control it returns to the first. */
	private trapTab(event: KeyboardEvent): void {
		const stops = this.tabStops();
		if (stops.length === 0) return;
		const root = this.getRootNode() as Document | ShadowRoot;
		const current = this.shadowRoot?.activeElement ?? root.activeElement;
		const index = stops.findIndex((control) => control === current || control.contains(current));
		const last = stops.length - 1;
		if (event.shiftKey ? index > 0 : index < last) return;
		event.preventDefault();
		stops[event.shiftKey ? last : 0]?.focus();
	}

	get selected(): string {
		return this.picked ?? '';
	}
	set selected(id: string | null) {
		this.picked = id;
		this.reflectAttribute('selected', id, {});
		this.requestUpdate('selected');
	}

	/** Open on a page (default: the selected one) and focus its item. */
	show(page = this.selected || this.items.find((item) => this.page(item.id))?.id || ''): void {
		this.open = true;
		if (page) this.selected = page;
		this.button(this.selected)?.focus();
	}

	/** Close without asking (emits `office-backstage-close` with reason `api`). */
	close(): void {
		if (!present(this.open)) return;
		this.open = false;
		this.fire('office-backstage-close', { reason: 'api' });
	}

	private page(id: string): HTMLElement | undefined {
		return [...this.children].find(
			(child): child is HTMLElement =>
				child instanceof HTMLElement && child.dataset.backstagePage === id,
		);
	}

	private button(id: string): HTMLButtonElement | null {
		return (
			[...this.renderRoot.querySelectorAll<HTMLButtonElement>('[data-backstage-item]')].find(
				(button) => button.dataset.backstageItem === id,
			) ?? null
		);
	}

	private choose(id: string): void {
		if (this.fire('office-backstage-select', { id }, true) && this.page(id)) this.selected = id;
	}

	private requestClose(reason: 'back' | 'escape'): void {
		if (this.fire('office-backstage-close', { reason }, true)) this.open = false;
	}

	private onNavClick(event: Event): void {
		const item = (event.target as Element).closest?.<HTMLButtonElement>('[data-backstage-item]');
		if (item && !item.disabled) this.choose(item.dataset.backstageItem!);
	}

	/** Exactly the selected page of the light DOM shows. */
	private syncPages(): void {
		for (const child of this.children)
			if (child instanceof HTMLElement && child.dataset.backstagePage !== undefined)
				child.hidden = child.dataset.backstagePage !== this.selected;
	}

	protected override updated(): void {
		this.syncPages();
	}

	private item(item: OfficeBackstageItem) {
		return html`<button
			type="button"
			class="item"
			data-backstage-item=${item.id}
			title=${ifDefined(item.title || undefined)}
			aria-current=${ifDefined(item.id === this.selected ? 'page' : undefined)}
			?disabled=${Boolean(item.disabled)}
			>${item.label}</button
		>`;
	}

	protected override render() {
		const label = this.label ?? 'File';
		const back = this.backLabel ?? 'Back';
		return html`
			<div class="frame" role="dialog" aria-modal="true" aria-label=${label}>
				<nav part="nav" aria-label=${label} @click=${this.onNavClick}>
					<button
						type="button"
						class="back"
						data-backstage="back"
						aria-label=${back}
						title="${back} (Esc)"
						@click=${() => this.requestClose('back')}
						>←</button
					>
					<div class="items"
						>${this.items.filter((item) => item.group !== 'footer').map((item) => this.item(item))}</div
					>
					<div class="footer"
						>${this.items.filter((item) => item.group === 'footer').map((item) => this.item(item))}</div
					>
				</nav>
				<div class="body" part="body"><slot @slotchange=${this.syncPages}></slot></div>
			</div>
		`;
	}
}

export const defineBackstage = definer('office-ui-backstage', () => OfficeUiBackstage);
