import { LitElement, html, nothing } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeElement, controlStyles, flag, triState } from '../base.js';
import { glyph } from '../glyph.js';
import { definer, present } from '../registry.js';
import css from './button.css?raw';

export type OfficeCommandEvent = CustomEvent<{ command: string }>;

type Configured = { requestEvent: string; idAttribute: string; detailKey: string };

const asAttribute = (value: boolean | undefined): string | undefined =>
	value === undefined ? undefined : String(value);

/**
 * Command button. Attributes: `label`, `icon`, `command`, `disabled`, `pressed` (aria-pressed),
 * `expanded` (aria-expanded), `icon-only`, `variant="stacked"`, `size="large|small"` (ribbon
 * large and small commands), `tall` (with `icon-only`: a tool tile), `active`, `badge`, `caret`
 * (a trailing menu chevron), `title`, `keyshortcuts` (forwarded as aria-keyshortcuts).
 * Activation (pointer, Enter, Space) emits one bubbling, composed `office-command` `{ command }`;
 * Enter and Space never reach the host's own key handlers. Nothing happens when disabled.
 * Static `requestEvent`, `idAttribute` and `detailKey`, and the `iconName()` and `showsCaret()`
 * methods, let a product subclass keep its published contract and glyph set.
 */
export class OfficeUiButton extends OfficeElement {
	static requestEvent = 'office-command';
	static idAttribute = 'command';
	static detailKey = 'command';
	static override shadowRootOptions = { ...LitElement.shadowRootOptions, delegatesFocus: true };
	static override styles = controlStyles(css);
	static override properties = { disabled: flag, pressed: triState };
	/** Everything else is an attribute, as it always was (see `OfficeElement.watched`). */
	static override watched = [
		'label',
		'icon',
		'command',
		'badge',
		'caret',
		'keyshortcuts',
		'size',
		'variant',
		'active',
		'tall',
		'icon-only',
		'expanded',
		'title',
	];
	declare disabled: boolean;
	declare pressed: boolean | undefined;

	constructor() {
		super();
		this.disabled = false;
	}

	/** The registered icon to draw; a product subclass may map names onto its glyph set. */
	protected iconName(): string | null {
		return this.getAttribute('icon');
	}
	/** The tooltip: `title`, else the label of an icon-only button; a product may always show it. */
	protected tooltip(label: string, iconOnly: boolean): string {
		return this.getAttribute('title') ?? (iconOnly ? label : '');
	}
	/** Whether the trailing menu chevron shows. */
	protected showsCaret(): boolean {
		const caret = this.getAttribute('caret');
		return caret !== null && caret !== 'false';
	}

	private onKeydown(event: KeyboardEvent): void {
		// Keep native activation out of the host's own shortcut handlers.
		if (
			(event.key === ' ' || event.key === 'Enter') &&
			!event.ctrlKey &&
			!event.metaKey &&
			!event.altKey
		)
			event.stopPropagation();
	}

	private onClick(): void {
		const { requestEvent, idAttribute, detailKey } = this.constructor as unknown as Configured;
		const id = this.getAttribute(idAttribute);
		if (!present(this.disabled) && id) this.fire(requestEvent, { [detailKey]: id });
	}

	protected override render() {
		const label = this.getAttribute('label') ?? '';
		const iconOnly = this.hasAttribute('icon-only');
		const badge = this.getAttribute('badge') ?? '';
		const expanded = this.getAttribute('expanded');
		const tooltip = this.tooltip(label, iconOnly);
		// The content is written without whitespace between nodes, so `textContent` is the label.
		return html`<button
			part="button"
			type="button"
			?disabled=${present(this.disabled)}
			title=${tooltip || nothing}
			aria-label=${ifDefined(iconOnly && label ? label : undefined)}
			aria-keyshortcuts=${ifDefined(this.getAttribute('keyshortcuts') || undefined)}
			aria-pressed=${ifDefined(asAttribute(this.pressed))}
			aria-expanded=${ifDefined(expanded === null ? undefined : String(expanded !== 'false'))}
			@click=${this.onClick}
			@keydown=${this.onKeydown}
			>${glyph(this.iconName())}<span class="label" ?hidden=${iconOnly}
				>${label}${glyph('chevronDown', 'caret', !this.showsCaret())}</span
			><span class="badge" aria-hidden="true" ?hidden=${!badge}>${badge}</span></button
		>`;
	}
}

export const defineButton = definer('office-ui-button', () => OfficeUiButton);
