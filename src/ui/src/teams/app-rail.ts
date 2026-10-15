import { html, nothing } from 'lit';
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
}

/**
 * The narrow app bar on the far left (Chat, Teams, Calls, Files); below 56em a bottom navigation bar. Properties: `items`,
 * `selected`, `hideLabels` (default false). Hidden labels retain accessible names and tooltips.
 * Emits `office-rail-select` `{ id }`.
 */
export class OfficeUiAppRail extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		items: { attribute: false },
		selected: { type: String },
		hideLabels: { type: Boolean, attribute: 'hide-labels', reflect: true },
	};
	declare items: RailItem[];
	declare selected: string;
	declare hideLabels: boolean;

	constructor() {
		super();
		this.items = [];
		this.selected = '';
		this.hideLabels = false;
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.setAttribute('role', 'navigation');
		this.setAttribute('aria-label', 'App bar');
	}

	protected override render() {
		return html`
			<ul>
				${this.items.map(
					(item) => html`
						<li>
							<button
								type="button"
								aria-label=${item.label}
								title=${item.label}
								aria-current=${this.selected === item.id ? 'page' : 'false'}
								@click=${() => this.fire('office-rail-select', { id: item.id })}
							>
								<span class="glyph">
									${icon(item.icon)}
									${item.badge ? html`<span class="badge">${item.badge > 99 ? '99+' : item.badge}</span>` : ''}
								</span>
								<span class=${item.shortLabel ? 'label has-short' : 'label'}>${item.label}</span>
								${item.shortLabel ? html`<span class="short" aria-hidden="true">${item.shortLabel}</span>` : nothing}
							</button>
						</li>
					`,
				)}
			</ul>
		`;
	}
}

export const defineAppRail = definer('office-ui-app-rail', () => OfficeUiAppRail);
