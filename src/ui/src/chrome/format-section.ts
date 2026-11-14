import { html, nothing } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { definer } from '../registry';
import css from './format-section.css?raw';

/**
 * `<office-ui-format-section label open mixed>`: one collapsible section of a Format pane
 * (Fill, Outline, Size and position, ...). A native disclosure (`<details>`/`<summary>`) in the
 * shadow root, so Enter and Space toggle it and assistive technology announces it as expanded or
 * collapsed; the slotted body stays mounted while closed. `open` is owned by the host: a click
 * on the heading flips it and emits `office-section-toggle` with `{ open }` (cancelable; a
 * cancelled event leaves the section as it was). `mixed` shows a small badge (`mixed-label`,
 * default "Mixed") for a multi-selection whose values differ. `focus()` focuses the heading.
 * Parts: `heading`, `label`, `badge` and `body`.
 */
export class OfficeUiFormatSection extends OfficeElement {
	static toggleEvent = 'office-section-toggle';
	static override styles = controlStyles(css);
	static override properties = {
		label: { type: String },
		open: flag,
		mixed: flag,
		mixedLabel: { type: String, attribute: 'mixed-label' },
	};
	declare label: string;
	declare open: boolean;
	declare mixed: boolean;
	declare mixedLabel: string;

	constructor() {
		super();
		this.label = '';
		this.open = false;
		this.mixed = false;
		this.mixedLabel = '';
	}

	override focus(options?: FocusOptions): void {
		this.shadowRoot?.querySelector<HTMLElement>('summary')?.focus(options);
	}

	private onToggle(event: Event): void {
		// The host owns `open`: stop the native toggle and flip it here, synchronously.
		event.preventDefault();
		const open = !this.open;
		const type = (this.constructor as typeof OfficeUiFormatSection).toggleEvent;
		if (this.fire(type, { open }, true)) this.open = open;
	}

	protected override render() {
		const mixed = this.mixed
			? html`<span class="badge" part="badge">${this.mixedLabel || 'Mixed'}</span>`
			: nothing;
		return html`<details ?open=${this.open}
			><summary part="heading" @click=${(event: Event) => this.onToggle(event)}
				>${glyph('chevronRight', 'chevron', true)}<span class="label" part="label"
					>${this.label}</span
				>${mixed}</summary
			><div class="body" part="body"><slot></slot></div
		></details>`;
	}
}

export const defineFormatSection = definer('office-ui-format-section', () => OfficeUiFormatSection);
