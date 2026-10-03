import { tok } from './tokens.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/**
 * Radio button with a native-like contract. Moved from pptx-viewer `pptx-ui-radio`
 * (`packages/shared/src/web-components/radio.ts`). Radios sharing a `name` in one tree and form
 * are a WAI-ARIA radiogroup with one roving tab stop (the checked radio, else the first enabled
 * one). Arrows move, select and focus the next enabled radio (wrapping), Home and End jump to the
 * ends, Space selects. Checking programmatically silently unchecks peers; only user activation
 * emits `input` then `change`, on the radio that becomes checked. Disabled radios are inert.
 * Tokens: `--office-checkbox-size` (16px), `--office-checkbox-size-touch` (22px),
 * `--office-radio-dot-size` (8px), `--office-radio-dot-size-touch` (12px).
 */
const CSS = `
:host { display: inline-grid; box-sizing: border-box; width: ${tok('--office-checkbox-size')};
	height: ${tok('--office-checkbox-size')}; flex: none; place-items: center;
	border: ${tok('--office-border-width')} solid ${tok('--office-checkbox-border')}; border-radius: 50%;
	background: ${tok('--office-checkbox-background')}; cursor: pointer; vertical-align: middle; }
:host([checked]) { border-color: ${tok('--office-checkbox-accent')}; }
:host(:focus-visible) { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
:host([disabled]) { opacity: .5; cursor: not-allowed; }
.dot { display: none; width: ${tok('--office-radio-dot-size')}; height: ${tok('--office-radio-dot-size')};
	border-radius: 50%; background: ${tok('--office-checkbox-accent')}; }
:host([checked]) .dot { display: block; }
@media (pointer: coarse), (max-width: 767px) {
	:host { width: ${tok('--office-checkbox-size-touch')}; height: ${tok('--office-checkbox-size-touch')}; }
	.dot { width: ${tok('--office-radio-dot-size-touch')}; height: ${tok('--office-radio-dot-size-touch')}; }
}
@media (forced-colors: active) {
	:host { border-color: CanvasText; background: Canvas; forced-color-adjust: auto; }
	:host([checked]) { border-color: Highlight; }
	.dot { background: Highlight; }
	:host([disabled]) { border-color: GrayText; }
	:host([disabled]) .dot { background: GrayText; }
	:host(:focus-visible) { outline-color: Highlight; }
}
`;

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
	const enabled = group.filter((item) => !item.disabled);
	const stop = enabled.find((item) => item.checked) ?? enabled[0];
	for (const item of group) item.tabIndex = item === stop ? 0 : -1;
}

export const defineRadio = definer('office-ui-radio', () => {
	class OfficeUiRadio extends HTMLElement {
		static formAssociated = true;
		static observedAttributes = ['checked', 'disabled', 'value', 'name'];
		private readonly internals: ElementInternals | undefined;
		private defaultChecked = false;
		private scope: ParentNode | null = null;
		private autoNamed = false;
		constructor() {
			super();
			try {
				this.internals = this.attachInternals();
			} catch {
				this.internals = undefined;
			}
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const dot = this.ownerDocument.createElement('span');
			dot.className = 'dot';
			root.append(dot);
			this.addEventListener('click', (event) => {
				if (this.disabled) {
					event.preventDefault();
					return;
				}
				this.select(true);
			});
			this.addEventListener('keydown', (event) => this.onKey(event));
			this.addEventListener('focus', () => this.syncName());
		}
		connectedCallback(): void {
			this.defaultChecked = this.checked;
			const root = this.getRootNode();
			this.scope = root instanceof Document || root instanceof ShadowRoot ? root : null;
			this.sync();
			this.syncName();
		}
		disconnectedCallback(): void {
			if (this.scope)
				refreshTabStops(groupOf(this as RadioLike, this.scope).filter((item) => item !== this));
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
		attributeChangedCallback(name: string): void {
			if (name === 'checked' && this.checked) this.uncheckPeers();
			this.sync();
		}
		get checked(): boolean {
			return this.hasAttribute('checked');
		}
		set checked(next: boolean) {
			this.toggleAttribute('checked', Boolean(next));
		}
		get disabled(): boolean {
			return this.hasAttribute('disabled');
		}
		set disabled(next: boolean) {
			this.toggleAttribute('disabled', Boolean(next));
		}
		get value(): string {
			return this.getAttribute('value') ?? 'on';
		}
		set value(next: string) {
			this.setAttribute('value', String(next));
		}
		get name(): string {
			return this.getAttribute('name') ?? '';
		}
		set name(next: string) {
			this.setAttribute('name', String(next));
		}
		get form(): HTMLFormElement | null {
			return this.internals?.form ?? null;
		}
		formResetCallback(): void {
			this.checked = this.defaultChecked;
		}
		formDisabledCallback(disabled: boolean): void {
			this.disabled = disabled;
		}
		/** Check this radio and, for user intent, announce it. No-op when already checked. */
		select(emit: boolean): void {
			if (this.checked) return;
			this.checked = true;
			if (emit) {
				this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
				this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
			}
		}
		private peers(): RadioLike[] {
			return groupOf(this as RadioLike, this.scope ?? (this.getRootNode() as ParentNode));
		}
		private uncheckPeers(): void {
			for (const peer of this.peers()) if (peer !== this && peer.checked) peer.checked = false;
		}
		private onKey(event: KeyboardEvent): void {
			if (this.disabled || event.ctrlKey || event.metaKey || event.altKey) return;
			if (event.key === ' ') {
				event.preventDefault();
				this.select(true);
				return;
			}
			const list = this.peers().filter((item) => !item.disabled);
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
		private sync(): void {
			this.setAttribute('role', 'radio');
			this.setAttribute('aria-checked', String(this.checked));
			this.setAttribute('aria-disabled', String(this.disabled));
			this.internals?.setFormValue?.(this.checked && !this.disabled ? this.value : null);
			if (this.isConnected) {
				this.scope = this.getRootNode() as ParentNode;
				refreshTabStops(this.peers());
			} else this.tabIndex = this.disabled ? -1 : 0;
		}
	}
	return OfficeUiRadio;
});
