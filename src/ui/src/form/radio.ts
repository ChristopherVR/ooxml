import { html } from 'lit';
import { OfficeElement, controlStyles, flag, valueOn } from '../base.js';
import { definer, present } from '../registry.js';
import css from './radio.css?raw';

interface RadioLike extends HTMLElement {
	checked: boolean;
	disabled: boolean;
	name: string;
	readonly form: HTMLFormElement | null;
	select(emit: boolean): void;
}

/** Radios of the same element type sharing a `name` in one tree and form form one group. */
function groupOf(radio: RadioLike, scope: ParentNode | null): RadioLike[] {
	const name = radio.name;
	if (!name || !scope) return [radio];
	return [...scope.querySelectorAll<RadioLike>(radio.localName)].filter(
		(other) => other.name === name && other.form === radio.form,
	);
}

/** Text of a wrapping label, ignoring other controls nested in it (a select, a field). */
function labelText(label: Element, self: Element): string {
	let text = '';
	for (const node of label.childNodes) {
		if (node === self) continue;
		if (node.nodeType === 3) text += node.textContent ?? '';
		else if (
			node instanceof Element &&
			!node.localName.includes('-') &&
			!/^(select|input|textarea)$/u.test(node.localName)
		)
			text += labelText(node, self);
	}
	return text;
}

/** One tab stop per group: the checked radio, else the first enabled one. */
function refreshTabStops(group: readonly RadioLike[]): void {
	const enabled = group.filter((item) => !present(item.disabled));
	const stop = enabled.find((item) => present(item.checked)) ?? enabled[0];
	for (const item of group) item.tabIndex = item === stop ? 0 : -1;
}

/**
 * Radio button with a native-like contract. Moved from pptx-viewer `pptx-ui-radio`. Radios
 * sharing a `name` in one tree and form are a WAI-ARIA radiogroup with one roving tab stop (the
 * checked radio, else the first enabled one). Arrows move, select and focus the next enabled
 * radio (wrapping), Home and End jump to the ends, Space selects. Checking programmatically
 * silently unchecks peers; only user activation emits `input` then `change`, on the radio that
 * becomes checked. Disabled radios are inert.
 */
export class OfficeUiRadio extends OfficeElement {
	static formAssociated = true;
	static override styles = controlStyles(css);
	static override properties = {
		checked: flag,
		disabled: flag,
		value: valueOn,
		name: { type: String, reflect: true },
	};
	declare checked: boolean;
	declare disabled: boolean;
	declare value: string;
	declare name: string;

	private readonly internals: ElementInternals | undefined;
	private defaultChecked = false;
	private scope: ParentNode | null = null;
	private autoNamed = false;

	constructor() {
		super();
		this.checked = false;
		this.disabled = false;
		this.value = 'on';
		this.name = '';
		try {
			this.internals = this.attachInternals();
		} catch {
			this.internals = undefined;
		}
		this.addEventListener('click', (event) => {
			if (present(this.disabled)) event.preventDefault();
			else this.select(true);
		});
		this.addEventListener('keydown', (event) => this.onKey(event));
		this.addEventListener('focus', () => this.syncName());
	}

	get form(): HTMLFormElement | null {
		return this.internals?.form ?? null;
	}

	override connectedCallback(): void {
		this.defaultChecked = present(this.checked);
		const root = this.getRootNode();
		this.scope = root instanceof Document || root instanceof ShadowRoot ? root : null;
		super.connectedCallback();
		this.syncName();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		if (this.scope)
			refreshTabStops(groupOf(this as RadioLike, this.scope).filter((item) => item !== this));
	}

	formResetCallback(): void {
		this.checked = this.defaultChecked;
	}

	formDisabledCallback(disabled: boolean): void {
		this.disabled = disabled;
	}

	/** Check this radio and, for user intent, announce it. No-op when already checked. */
	select(emit: boolean): void {
		if (present(this.checked)) return;
		this.checked = true;
		if (emit) {
			this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
			this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
		}
	}

	/**
	 * A custom element is not named by a wrapping `<label>` in every browser's accessibility
	 * tree, so mirror the label text into `aria-label` unless the host named it.
	 */
	private syncName(): void {
		if (this.hasAttribute('aria-labelledby')) return;
		if (this.hasAttribute('aria-label') && !this.autoNamed) return;
		const label = this.closest('label') ?? (this.internals?.labels?.[0] as Element | undefined);
		const text = label ? labelText(label, this).replace(/\s+/gu, ' ').trim() : '';
		if (text) {
			this.autoNamed = true;
			this.setAttribute('aria-label', text);
		}
	}

	private peers(): RadioLike[] {
		return groupOf(this as RadioLike, this.scope ?? (this.getRootNode() as ParentNode));
	}

	private onKey(event: KeyboardEvent): void {
		if (present(this.disabled) || event.ctrlKey || event.metaKey || event.altKey) return;
		if (event.key === ' ') {
			event.preventDefault();
			this.select(true);
			return;
		}
		const list = this.peers().filter((item) => !present(item.disabled));
		const index = list.indexOf(this as RadioLike);
		const last = list.length - 1;
		const next =
			event.key === 'ArrowRight' || event.key === 'ArrowDown'
				? (index + 1) % list.length
				: event.key === 'ArrowLeft' || event.key === 'ArrowUp'
					? (index + last) % list.length
					: event.key === 'Home'
						? 0
						: event.key === 'End'
							? last
							: -1;
		if (next < 0 || index < 0) return;
		event.preventDefault();
		event.stopPropagation();
		const target = list[next]!;
		target.select(true);
		target.focus();
	}

	/** The host carries the semantics: role, state, form value and the group's tab stop. */
	protected override willUpdate(): void {
		const checked = present(this.checked);
		const disabled = present(this.disabled);
		this.setAttribute('role', 'radio');
		this.setAttribute('aria-checked', String(checked));
		this.setAttribute('aria-disabled', String(disabled));
		this.internals?.setFormValue?.(checked && !disabled ? this.value : null);
		if (this.isConnected) {
			this.scope = this.getRootNode() as ParentNode;
			if (checked)
				for (const peer of this.peers())
					if (peer !== this && present(peer.checked)) peer.checked = false;
			refreshTabStops(this.peers());
		} else this.setAttribute('tabindex', disabled ? '-1' : '0');
	}

	protected override render() {
		return html`<span class="dot"></span>`;
	}
}

export const defineRadio = definer('office-ui-radio', () => OfficeUiRadio);
