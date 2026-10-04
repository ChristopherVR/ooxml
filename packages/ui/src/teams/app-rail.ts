import { html } from 'lit';
import { definer } from '../registry.js';
import { TeamsElement, withStyles } from './base.js';
import { icon } from './icons.js';
import css from './app-rail.css?raw';

export interface RailItem {
	id: string;
	label: string;
	/** An icon name from the ooxml-ui icon registry (`chat`, `users`, `phone`, `folder`, ...). */
	icon: string;
	badge?: number;
}

/**
 * The narrow app bar on the far left (Chat, Teams, Calls, Files). Properties: `items`,
 * `selected`. Emits `office-rail-select` `{ id }`.
 */
export class OfficeUiAppRail extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		items: { attribute: false },
		selected: { type: String },
	};
	declare items: RailItem[];
	declare selected: string;

	constructor() {
		super();
		this.items = [];
		this.selected = '';
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
								aria-current=${this.selected === item.id ? 'page' : 'false'}
								@click=${() => this.fire('office-rail-select', { id: item.id })}
							>
								<span class="glyph">
									${icon(item.icon)}
									${item.badge ? html`<span class="badge">${item.badge > 99 ? '99+' : item.badge}</span>` : ''}
								</span>
								<span class="label">${item.label}</span>
							</button>
						</li>
					`,
				)}
			</ul>
		`;
	}
}

export const defineAppRail = definer('office-ui-app-rail', () => OfficeUiAppRail);
