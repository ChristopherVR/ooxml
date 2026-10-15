import { html, type PropertyValues, type TemplateResult } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeElement, controlStyles, flag } from '../base';
import { defineDialog } from './dialog';
import { definer, present } from '../registry';
import {
	optionRow,
	type OfficeOptionControl,
	type OfficeOptionValue,
	type OfficeOptionValues,
} from './options-controls';
import css from './options-dialog.css?raw';

/**
 * Office File > Options dialog. Products describe their categories as data; the element renders
 * the category list, the pane and Reset/Cancel/OK, and reports the edited values. Generalised from
 * pptx-viewer's `VIEWER_OPTIONS_SCHEMA` dialog (render/options) so every viewer shares one, in the
 * same arrangement: a category rail, a headline, small uppercase section titles and rows with the
 * label at the start and the control at the end.
 */
export {
	clampOptionNumber,
	type OfficeOptionChoice,
	type OfficeOptionControl,
	type OfficeOptionValue,
	type OfficeOptionValues,
} from './options-controls';

export interface OfficeOptionSection {
	id: string;
	title: string;
	description?: string;
	controls: readonly OfficeOptionControl[];
	/**
	 * The product supplies a block of its own under this section's rows (PowerPoint's theme
	 * picker): put an element with `slot="<this value>"` in the dialog and it shows there while
	 * the category is open.
	 */
	slot?: string;
}

export interface OfficeOptionCategory {
	id: string;
	label: string;
	/** Headline at the top of the pane. */
	description?: string;
	sections: readonly OfficeOptionSection[];
	/** The whole category is not available yet; the pane explains why. */
	disabled?: string;
	/**
	 * The product supplies this pane itself (Customize Ribbon): put an element with
	 * `slot="pane-<id>"` in the dialog and it shows under the sections while the category is open.
	 */
	custom?: boolean;
}

/** Translatable chrome of the dialog (English when unset). */
export interface OfficeOptionsLabels {
	ok: string;
	cancel: string;
	reset: string;
	categories: string;
	close: string;
}
const DEFAULT_LABELS: OfficeOptionsLabels = {
	ok: 'OK',
	cancel: 'Cancel',
	reset: 'Reset all options',
	categories: 'Categories',
	close: 'Close',
};

/** Emitted on OK with every value and the keys that changed. */
export type OfficeOptionsChangeEvent = CustomEvent<{
	values: OfficeOptionValues;
	changed: string[];
}>;

/**
 * `<office-ui-options-dialog>`: set `categories` and `values`, then `show()` (or the `open`
 * attribute). Edits stay in a draft until OK, which emits `office-options-change` and closes;
 * Cancel, Escape and the close button discard them and emit `office-options-cancel`. "Reset all options" (shown when `defaults` is
 * set) puts the defaults in the draft. `heading` defaults to "Options"; `labels` translates the
 * buttons.
 *
 * `live` is PowerPoint's behaviour: every edit applies at once (there is no draft), OK, Escape
 * and the close button keep the edits and Cancel emits `office-options-cancel` for the host to
 * restore its snapshot. "Reset all options" then only emits `office-options-reset`; the host
 * resets and sets `values` again. Live or not, every edit emits `office-option-input` (key and
 * value) and selecting a category emits `office-options-category`.
 *
 * `custom-body` hands the whole body to the host: the frame, heading and close button stay, and
 * whatever the host puts in `slot="body"` replaces the category rail, the pane and the footer.
 * The close button and Escape then behave as they do without it (OK when `live`, else Cancel).
 */
export class OfficeUiOptionsDialog extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		open: flag,
		live: flag,
		customBody: { ...flag, attribute: 'custom-body' },
		heading: { type: String },
		categories: { attribute: false },
		defaults: { attribute: false },
		labels: { attribute: false },
		// `values` keeps a private copy and `category` validates its id, so both have accessors.
		active: { state: true },
	};
	declare open: boolean;
	declare live: boolean;
	declare customBody: boolean;
	declare heading: string | null;
	declare categories: readonly OfficeOptionCategory[];
	declare defaults: OfficeOptionValues | null;
	declare labels: Partial<OfficeOptionsLabels>;
	declare active: string;
	private committed: OfficeOptionValues = {};
	private draftValues: OfficeOptionValues = {};

	constructor() {
		super();
		this.open = false;
		this.live = false;
		this.customBody = false;
		this.heading = null;
		this.categories = [];
		this.defaults = null;
		this.labels = {};
		this.active = '';
	}

	get draft(): OfficeOptionValues {
		return this.draftValues;
	}

	get values(): OfficeOptionValues {
		return { ...this.committed };
	}
	set values(value: OfficeOptionValues) {
		this.committed = { ...value };
		if (!present(this.open) || present(this.live)) this.draftValues = { ...value };
		this.requestUpdate('values');
	}

	/** The selected category id. */
	get category(): string {
		return this.active;
	}
	set category(id: string) {
		if (this.categories.some((category) => category.id === id)) this.active = id;
	}

	show(): void {
		this.open = true;
	}

	/** Closes without applying: the draft is dropped and `office-options-cancel` is emitted. */
	close(): void {
		const wasOpen = present(this.open);
		this.open = false;
		if (wasOpen) this.fire('office-options-cancel', {});
	}

	private commit(): void {
		const changed = Object.keys(this.draftValues).filter(
			(key) => this.draftValues[key] !== this.committed[key],
		);
		this.committed = { ...this.draftValues };
		this.fire('office-options-change', { values: { ...this.committed }, changed });
		this.open = false;
	}

	private resetAll(): void {
		if (present(this.live)) {
			this.fire('office-options-reset', {});
			return;
		}
		this.draftValues = { ...this.draftValues, ...this.defaults };
		this.requestUpdate();
	}

	/** Panes the product slots in sit outside the dialog's own focus trap: let Tab walk through them. */
	private onPaneKey(event: KeyboardEvent): void {
		if (event.key !== 'Tab') return;
		const origin = event.composedPath()[0];
		if (origin instanceof Node && !this.renderRoot.contains(origin)) event.stopPropagation();
	}

	private onNavKey(event: KeyboardEvent): void {
		const tabs = [...this.renderRoot.querySelectorAll<HTMLButtonElement>('nav [role="tab"]')];
		const at = tabs.findIndex((tab) => tab.getAttribute('aria-selected') === 'true');
		const step = { ArrowDown: 1, ArrowRight: 1, ArrowUp: -1, ArrowLeft: -1 }[event.key];
		const target =
			event.key === 'Home'
				? tabs[0]
				: event.key === 'End'
					? tabs.at(-1)
					: step
						? tabs[(at + step + tabs.length) % tabs.length]
						: undefined;
		if (!target) return;
		event.preventDefault();
		this.category = target.dataset.category!;
		this.renderRoot.querySelector<HTMLButtonElement>('nav [aria-selected="true"]')?.focus();
	}

	/** Native controls and popups follow the theme: dark when `--office-background` is dark. */
	private syncColorScheme(): void {
		const probe = this.ownerDocument.createElement('span');
		probe.style.cssText = 'display:none;background-color:var(--office-background)';
		this.renderRoot.append(probe);
		const match = /\d+(\.\d+)?/g;
		let background = '';
		try {
			background = getComputedStyle(probe).backgroundColor;
		} catch {
			// A host whose token chain cannot be resolved keeps the default scheme.
		}
		probe.remove();
		if (!background) return;
		const [r = 255, g = 255, b = 255] = (background.match(match) ?? []).slice(0, 3).map(Number);
		this.style.colorScheme = 0.2126 * r + 0.7152 * g + 0.0722 * b < 128 ? 'dark' : 'light';
	}

	protected override updated(changed: PropertyValues<this>): void {
		if (changed.has('open') && present(this.open)) this.syncColorScheme();
		if (changed.has('active') && this.active)
			this.fire('office-options-category', { id: this.active });
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (!this.categories.some((category) => category.id === this.active))
			this.active = this.categories[0]?.id ?? '';
		// Opening starts a fresh draft from the committed values.
		if (changed.has('open') && present(this.open)) this.draftValues = { ...this.committed };
	}

	set(key: string, next: OfficeOptionValue): void {
		this.draftValues[key] = next;
		if (present(this.live)) this.committed[key] = next;
		this.fire('office-option-input', { key, value: next });
		// Other rows may depend on this value (`enabledBy`).
		this.requestUpdate();
	}

	private pane(category: OfficeOptionCategory | undefined): TemplateResult | '' {
		if (!category) return '';
		return html`
			${category.description ? html`<p class="lead">${category.description}</p>` : ''}
			${category.disabled ? html`<p class="unavailable">${category.disabled}</p>` : ''}
			${category.sections.map(
				(section) =>
					html`<section>
						<h3>${section.title}</h3>
						${section.description ? html`<p class="desc">${section.description}</p>` : ''}
						${section.controls.map((control) => optionRow(this, control))}
						${section.slot ? html`<slot name=${section.slot}></slot>` : ''}
					</section>`,
			)}
			${category.custom ? html`<slot name="pane-${category.id}"></slot>` : ''}
		`;
	}

	protected override render() {
		const category = this.categories.find((entry) => entry.id === this.active);
		const text = { ...DEFAULT_LABELS, ...this.labels };
		return html`
			<office-ui-dialog
				heading=${this.heading ?? 'Options'}
				close-label=${text.close}
				?open=${present(this.open)}
				@office-dialog-close=${(event: Event) => {
					if (event.defaultPrevented) return;
					if (present(this.live)) this.commit();
					else this.close();
				}}
			>
				${
					present(this.customBody)
						? html`<div class="custom" @keydown=${this.onPaneKey}><slot name="body"></slot></div>`
						: this.layout(category, text)
				}
			</office-ui-dialog>
		`;
	}

	private layout(
		category: OfficeOptionCategory | undefined,
		text: OfficeOptionsLabels,
	): TemplateResult {
		return html`
			<div class="layout">
				<nav
					role="tablist"
					aria-orientation="vertical"
					aria-label=${text.categories}
					@keydown=${this.onNavKey}
				>
					${this.categories.map(
						(entry) =>
							html`<button
								type="button"
								role="tab"
								id="tab-${entry.id}"
								data-category=${entry.id}
								aria-controls="pane"
								aria-selected=${String(entry.id === this.active)}
								tabindex=${entry.id === this.active ? 0 : -1}
								@click=${() => (this.category = entry.id)}
								>${entry.label}</button
							>`,
					)}
				</nav>
				<div
					role="tabpanel"
					id="pane"
					@keydown=${this.onPaneKey}
					aria-labelledby=${ifDefined(category ? `tab-${category.id}` : undefined)}
					>${this.pane(category)}</div
				>
			</div>
			<div slot="footer" class="actions">
				${
					this.defaults
						? html`<button type="button" class="reset" data-action="reset" @click=${this.resetAll}>
								${text.reset}
							</button>`
						: ''
				}
				<button type="button" data-action="cancel" @click=${this.close}>${text.cancel}</button>
				<button type="button" class="primary" data-action="ok" @click=${this.commit}>
					${text.ok}
				</button>
			</div>
		`;
	}
}

export const defineOptionsDialog = definer(
	'office-ui-options-dialog',
	() => OfficeUiOptionsDialog,
	[defineDialog],
);
