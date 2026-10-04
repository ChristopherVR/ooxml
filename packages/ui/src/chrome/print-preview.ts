import { html, type PropertyValues } from 'lit';
import { OfficeElement, controlStyles } from '../base.js';
import { definer } from '../registry.js';
import css from './print-preview.css?raw';

export type OfficePrintPreviewPageEvent = CustomEvent<{ index: number }>;

/**
 * `<office-ui-print-preview>`: the Print page's paper preview with page navigation. Set `pages`
 * to rendered nodes (for example the SVG of each page); the element shows a deep clone of the
 * current one with `tabindex` and inline `style` removed, so the preview is inert and never
 * parses markup. `index` selects the page; the arrows emit `office-print-preview-page`.
 * Attributes: `label` (default "Print preview"), `empty-label`.
 */
export class OfficeUiPrintPreview extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		// `pages` clamps the index and `index` clamps itself to the pages, so both have accessors.
		pages: { attribute: false, noAccessor: true },
		index: { attribute: false, noAccessor: true },
		label: { type: String },
		emptyLabel: { attribute: 'empty-label', type: String },
	};
	declare label: string | null;
	declare emptyLabel: string | null;
	private list: readonly Node[] = [];
	private current = 0;
	private clone: Node | undefined;

	constructor() {
		super();
		this.label = null;
		this.emptyLabel = null;
	}

	get pages(): readonly Node[] {
		return this.list;
	}
	set pages(value: readonly Node[]) {
		this.list = value;
		this.current = Math.min(this.current, Math.max(0, value.length - 1));
		this.requestUpdate('pages');
	}

	get index(): number {
		return this.current;
	}
	set index(value: number) {
		const next = Math.max(0, Math.min(this.list.length - 1, Math.trunc(value) || 0));
		if (next === this.current) return;
		this.current = next;
		this.requestUpdate('index');
	}

	private go(index: number): void {
		const before = this.current;
		this.index = index;
		if (this.current !== before) this.fire('office-print-preview-page', { index: this.current });
	}

	/** The shown page is an inert copy: no tabindex, no inline style. */
	protected override willUpdate(changed: PropertyValues<this>): void {
		if (
			!changed.has('pages' as never) &&
			!changed.has('index' as never) &&
			this.clone !== undefined
		)
			return;
		const page = this.list[this.current];
		if (!page) {
			this.clone = undefined;
			return;
		}
		const clone = this.ownerDocument.importNode(page, true);
		if (clone instanceof Element) {
			clone.removeAttribute('style');
			clone.removeAttribute('tabindex');
			clone.querySelectorAll('[tabindex]').forEach((node) => node.removeAttribute('tabindex'));
		}
		this.clone = clone;
	}

	protected override render() {
		const total = this.list.length;
		const label = this.label ?? 'Print preview';
		return html`
			<div
				class="sheet"
				role="img"
				aria-label=${total > 1 ? `${label}, page ${this.current + 1} of ${total}` : label}
				>${this.clone ?? html`<p class="empty">${this.emptyLabel ?? 'Nothing to print.'}</p>`}</div
			>
			<div class="nav" ?hidden=${total < 2}>
				<button
					type="button"
					aria-label="Previous page"
					?disabled=${this.current === 0}
					@click=${() => this.go(this.current - 1)}
					>‹</button
				>
				<span>${this.current + 1} of ${total}</span>
				<button
					type="button"
					aria-label="Next page"
					?disabled=${this.current >= total - 1}
					@click=${() => this.go(this.current + 1)}
					>›</button
				>
			</div>
		`;
	}
}

export const definePrintPreview = definer('office-ui-print-preview', () => OfficeUiPrintPreview);
