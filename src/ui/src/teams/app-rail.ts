import { html, nothing, type PropertyValues } from 'lit';
import { definer } from '../registry';
import { TeamsElement, withStyles } from './base';
import { icon } from './icons';
import css from './app-rail.css?raw';

export interface RailItem {
	id: string;
	label: string;
	/** An icon name from the ooxml-ui icon registry (`chat`, `users`, `phone`, `folder`, ...). */
	icon: string;
	badge?: number;
	/** Label for the compact bottom bar, where space is tight. */
	shortLabel?: string;
	/** On the phone bottom bar this item lives in the More sheet instead of the bar. */
	secondary?: boolean;
}

const total = (n: number): string => (n > 99 ? '99+' : String(n));

/**
 * The narrow app bar on the far left (Chat, Teams, Calls, Files); below 56em a bottom navigation
 * bar showing the primary items plus More, which opens a sheet with the `secondary` ones (their
 * unread badges are carried to More). Properties: `items`, `selected`, `hideLabels` (default
 * false). Hidden labels retain accessible names and tooltips. Emits `office-rail-select` `{ id }`.
 */
export class OfficeUiAppRail extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		items: { attribute: false },
		selected: { type: String },
		hideLabels: { type: Boolean, attribute: 'hide-labels', reflect: true },
		moreOpen: { state: true },
	};
	declare items: RailItem[];
	declare selected: string;
	declare hideLabels: boolean;
	declare moreOpen: boolean;

	constructor() {
		super();
		this.items = [];
		this.selected = '';
		this.hideLabels = false;
		this.moreOpen = false;
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.setAttribute('role', 'navigation');
		this.setAttribute('aria-label', 'App bar');
	}

	protected override updated(changed: PropertyValues<this>): void {
		const dialog = this.renderRoot.querySelector('dialog');
		if (!dialog || !changed.has('moreOpen')) return;
		if (this.moreOpen && !dialog.open) dialog.showModal();
		else if (!this.moreOpen && dialog.open) dialog.close();
	}

	private closeMore(): void {
		this.moreOpen = false;
		// Hand focus back to the button that opened the sheet.
		void this.updateComplete.then(() =>
			this.renderRoot.querySelector<HTMLButtonElement>('.more')?.focus(),
		);
	}

	private select(id: string): void {
		this.moreOpen = false;
		this.fire('office-rail-select', { id });
	}

	protected override render() {
		const secondary = this.items.filter((i) => i.secondary);
		const badge = secondary.reduce((n, i) => n + (i.badge ?? 0), 0);
		const inMore = secondary.some((i) => i.id === this.selected);
		return html`
			<ul>
				${this.items.map(
					(item) => html`
						<li data-secondary=${String(item.secondary === true)}>${this.button(item)}</li>
					`,
				)}
				${
					secondary.length
						? html`<li class="more-item">
								<button
									type="button"
									class="more"
									aria-label=${badge ? `More, ${badge} unread` : 'More'}
									aria-haspopup="dialog"
									aria-expanded=${String(this.moreOpen)}
									aria-current=${inMore ? 'page' : 'false'}
									@click=${() => (this.moreOpen = true)}
								>
									<span class="glyph">
										${icon('more')}
										${badge ? html`<span class="badge">${total(badge)}</span>` : nothing}
									</span>
									<span class="short" aria-hidden="true">More</span>
								</button>
							</li>`
						: nothing
				}
			</ul>
			${
				secondary.length
					? html`<dialog
							class="sheet"
							aria-label="More apps"
							@cancel=${(e: Event) => {
								e.preventDefault();
								this.closeMore();
							}}
							@click=${(e: MouseEvent) => {
								if (e.target === e.currentTarget) this.closeMore();
							}}
						>
							<ul>
								${secondary.map((item) => html`<li>${this.button(item, true)}</li>`)}
							</ul>
							<button type="button" class="close" @click=${() => this.closeMore()}>Close</button>
						</dialog>`
					: nothing
			}
		`;
	}

	private button(item: RailItem, sheet = false) {
		return html`<button
			type="button"
			aria-label=${item.label}
			title=${item.label}
			aria-current=${this.selected === item.id ? 'page' : 'false'}
			@click=${() => this.select(item.id)}
		>
			<span class="glyph">
				${icon(item.icon)}
				${item.badge ? html`<span class="badge">${total(item.badge)}</span>` : nothing}
			</span>
			<span class=${item.shortLabel && !sheet ? 'label has-short' : 'label'}>${item.label}</span>
			${item.shortLabel && !sheet ? html`<span class="short" aria-hidden="true">${item.shortLabel}</span>` : nothing}
		</button>`;
	}
}

export const defineAppRail = definer('office-ui-app-rail', () => OfficeUiAppRail);
