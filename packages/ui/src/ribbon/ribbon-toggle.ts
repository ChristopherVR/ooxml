import { html, unsafeStatic } from 'lit/static-html.js';
import { OfficeElement, controlStyles, flag } from '../base.js';
import { definer, present } from '../registry.js';
import css from './ribbon-toggle.css?raw';

type Configured = { requestEvent: string; idAttribute: string; detailKey: string };
type CheckboxLike = HTMLElement & { checked: boolean };

/**
 * A ribbon checkbox row (View > Ruler, Gridlines): a labelled `office-ui-checkbox` sharing its
 * keyboard and focus behaviour. Moved from pptx-viewer `pptx-ui-ribbon-toggle`. Controlled:
 * `label`, `checked`, `disabled`, `title` and `command` attributes; user activation emits
 * `office-toggle` `{ command, checked }` (static `requestEvent`, `idAttribute` and `detailKey`
 * let a product subclass keep its published names). The host reflects the new state.
 */
export class OfficeUiRibbonToggle extends OfficeElement {
	static requestEvent = 'office-toggle';
	static idAttribute = 'command';
	static detailKey = 'command';
	/** The checkbox drawn inside; a product subclass may use its own aliased tag. */
	static checkboxTag = 'office-ui-checkbox';
	static override styles = controlStyles(css);
	static override properties = {
		label: { type: String },
		checked: flag,
		disabled: flag,
	};
	declare label: string;
	declare checked: boolean;
	declare disabled: boolean;

	/** `title` is a native attribute, so it is watched rather than declared as a property. */
	static override get observedAttributes(): string[] {
		return [...super.observedAttributes, 'title'];
	}

	constructor() {
		super();
		this.label = '';
		this.checked = false;
		this.disabled = false;
	}

	override attributeChangedCallback(name: string, old: string | null, value: string | null): void {
		super.attributeChangedCallback(name, old, value);
		if (name === 'title') this.requestUpdate();
	}

	/** Clicking the label text presses the checkbox, which is the only control that toggles. */
	private onLabelClick(event: Event): void {
		event.preventDefault();
		const checkbox = this.renderRoot.querySelector<HTMLElement>('[part="checkbox"]');
		if (event.target !== checkbox) checkbox?.click();
	}

	private onChange(event: Event): void {
		event.stopPropagation();
		const checkbox = event.target as CheckboxLike;
		const checked = checkbox.checked;
		// Controlled: show the host's state, not the checkbox's optimistic flip.
		checkbox.checked = present(this.checked);
		const { requestEvent, idAttribute, detailKey } = this.constructor as unknown as Configured;
		const id = this.getAttribute(idAttribute);
		if (!present(this.disabled) && id) this.fire(requestEvent, { [detailKey]: id, checked });
	}

	protected override render() {
		const checkboxTag = (this.constructor as unknown as { checkboxTag: string }).checkboxTag;
		const tag = unsafeStatic(checkboxTag);
		return html`
			<label @click=${this.onLabelClick}>
				<${tag}
					part="checkbox"
					aria-label=${this.label}
					title=${this.getAttribute('title') ?? this.label}
					?checked=${present(this.checked)}
					?disabled=${present(this.disabled)}
					@input=${(event: Event) => event.stopPropagation()}
					@change=${this.onChange}
				></${tag}>${this.label}</label>
		`;
	}
}

export const defineRibbonToggle = definer('office-ui-ribbon-toggle', () => OfficeUiRibbonToggle);
