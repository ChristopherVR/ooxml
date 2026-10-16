import { html, nothing } from 'lit';
import type { ContactView } from 'ooxml-core/teams';
import { definer } from '../registry';
import { TeamsElement, withStyles } from './base';
import css from './people-picker.css?raw';

/**
 * Pick people from the workspace roster: a search box and a list of checkable people with their
 * presence. Properties: `contacts` (core `ContactView[]`), `selected` (user ids), `exclude`
 * (user ids not offered, for example people already in the chat), `label`.
 * Emits `office-people-change` `{ selected }` on every toggle.
 */
export class OfficeUiPeoplePicker extends TeamsElement {
	static override styles = withStyles(css);
	static override properties = {
		contacts: { attribute: false },
		selected: { attribute: false },
		exclude: { attribute: false },
		label: { type: String },
		query: { state: true },
	};
	declare contacts: ContactView[];
	declare selected: readonly string[];
	declare exclude: readonly string[];
	declare label: string;
	declare query: string;

	constructor() {
		super();
		this.contacts = [];
		this.selected = [];
		this.exclude = [];
		this.label = '';
		this.query = '';
	}

	private toggle(id: string): void {
		const next = this.selected.includes(id)
			? this.selected.filter((s) => s !== id)
			: [...this.selected, id];
		this.fire('office-people-change', { selected: next });
	}

	protected override render() {
		const q = this.query.trim().toLocaleLowerCase();
		const offered = this.contacts.filter(
			(c) => !this.exclude.includes(c.id) && (!q || c.name.toLocaleLowerCase().includes(q)),
		);
		const chosen = this.contacts.filter((c) => this.selected.includes(c.id));
		const label = this.label || this.t('people.label');
		return html`
			${
				chosen.length
					? html`<ul class="chips" aria-label=${this.t('people.selected')}>
							${chosen.map(
								(c) =>
									html`<li>
										<button
											type="button"
											class="chip"
											aria-label=${this.t('people.remove', { name: c.name })}
											@click=${() => this.toggle(c.id)}
										>
											${c.name} <span aria-hidden="true">×</span>
										</button>
									</li>`,
							)}
						</ul>`
					: nothing
			}
			<input
				type="search"
				placeholder=${this.t('people.search')}
				aria-label=${this.t('people.search')}
				.value=${this.query}
				@input=${(e: Event) => (this.query = (e.target as HTMLInputElement).value)}
			/>
			${
				offered.length === 0
					? html`<p class="none">
							${this.contacts.length === 0 ? this.t('people.nobodyElse') : this.t('people.noMatch')}
						</p>`
					: html`<ul class="list" role="group" aria-label=${label}>
							${offered.map(
								(c) =>
									html`<li>
										<label>
											<input
												type="checkbox"
												.checked=${this.selected.includes(c.id)}
												@change=${() => this.toggle(c.id)}
											/>
											<office-ui-avatar
												size="sm"
												name=${c.name}
												seed=${c.id}
												presence=${c.availability}
											></office-ui-avatar>
											<span class="name">${c.name}</span>
										</label>
									</li>`,
							)}
						</ul>`
			}
		`;
	}
}

export const definePeoplePicker = definer('office-ui-people-picker', () => OfficeUiPeoplePicker);
