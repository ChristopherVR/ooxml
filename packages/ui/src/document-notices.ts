import { createIconSvg, paintIcon } from './icons.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/**
 * The "read-only recommended" banner and the Paste Options strip, moved from pptx-viewer
 * (`packages/shared/src/web-components/{read-only-banner,paste-options}.ts`). Word, Excel and
 * PowerPoint all recommend read-only opening and offer Paste Options after a paste. Text arrives
 * translated; event names and test ids are static class fields a product subclass may override.
 */

type Configured = { requestEvent: string; dismissEvent?: string; testIdPrefix: string };
const config = (el: HTMLElement): Configured => el.constructor as unknown as Configured;
const emit = (host: HTMLElement, type: string, detail?: unknown) =>
	host.dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));

// ---------------------------------------------------------------------------------------------
// Read-only banner
// ---------------------------------------------------------------------------------------------

export interface OfficeReadOnlyBannerState {
	/** Why read-only is recommended; mirrored onto the host's `data-kind`. */
	kind: string | null;
	/** Translated reason. */
	message: string;
	/** Swap Edit anyway / Dismiss for the inline password form. */
	passwordPromptOpen?: boolean;
	/** Translated password error, or null. */
	passwordError?: string | null;
	/** Disables the form while a submitted password is checked. */
	checkingPassword?: boolean;
	labels?: {
		title?: string;
		editAnyway?: string;
		dismiss?: string;
		passwordLabel?: string;
		passwordPlaceholder?: string;
		unlock?: string;
		cancel?: string;
	};
}

export type OfficeReadOnlyBannerIntent =
	| { id: 'editAnyway' | 'dismiss' | 'cancelPassword' }
	| { id: 'submitPassword'; password: string };

const BANNER_CSS = `
:host { display: block; flex: none; }
:host([hidden]) { display: none !important; }
[hidden] { display: none !important; }
svg { flex: none; width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.45;
	stroke-linecap: round; stroke-linejoin: round; color: #d97706; }
.banner { box-sizing: border-box; display: flex; align-items: center; gap: 12px; padding: 6px 16px;
	border-bottom: 1px solid rgba(180, 83, 9, .3); background: var(--office-notice-background, #fef3c7);
	color: var(--office-notice-foreground, #78350f); font: 12px/1.4 var(--office-font, system-ui, sans-serif); }
.text { flex: 1 1 auto; min-width: 0; margin: 0; }
.form { display: flex; flex: none; align-items: center; gap: 8px; }
.sr-only { position: absolute; width: 1px; height: 1px; margin: -1px; padding: 0; overflow: hidden;
	clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0; }
button, input { font: inherit; }
button { box-sizing: border-box; flex: none; min-height: 24px; padding: 4px 12px; border: 1px solid transparent;
	border-radius: 4px; background: transparent; color: inherit; font-weight: 500; cursor: pointer; touch-action: manipulation; }
button.primary { border-color: rgba(217, 119, 6, .5); }
button:hover { background: rgba(180, 83, 9, .15); }
button:disabled { opacity: .6; cursor: not-allowed; }
button:focus-visible, input:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 1px; }
input { box-sizing: border-box; min-height: 24px; padding: 4px 8px; border: 1px solid rgba(217, 119, 6, .4);
	border-radius: 4px; background: var(--office-background, #fff); color: var(--office-foreground, #1f2937); }
.error { flex: none; color: var(--office-danger, #b91c1c); }
@media (pointer: coarse) { button, input { min-height: 44px; min-width: 44px; } }
@media (forced-colors: active) {
	.banner { border-bottom-color: CanvasText; background: Canvas; color: CanvasText; }
	svg, .error { color: CanvasText; }
	button, input { border-color: ButtonText; color: ButtonText; background: ButtonFace; }
	button:focus-visible, input:focus-visible { outline-color: Highlight; }
}
`;

/**
 * `<office-ui-read-only-banner>`: controlled banner. The host owns the lock (Edit anyway,
 * Dismiss, the password check); the element owns the markup, the password form and its focus.
 * Each activation emits `office-read-only-request` with an {@link OfficeReadOnlyBannerIntent}.
 */
export const defineReadOnlyBanner = definer('office-ui-read-only-banner', () => {
	class OfficeUiReadOnlyBanner extends HTMLElement {
		static requestEvent = 'office-read-only-request';
		static testIdPrefix = 'office-readonly';
		private model: OfficeReadOnlyBannerState = { kind: null, message: '' };
		private promptWasOpen = false;
		private pendingFocus = false;
		private readonly parts = this.build();
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(BANNER_CSS));
			root.append(this.parts.banner);
		}
		get state(): OfficeReadOnlyBannerState {
			return this.model;
		}
		set state(value: OfficeReadOnlyBannerState) {
			this.model = value;
			this.render();
		}
		connectedCallback(): void {
			this.render();
			if (this.pendingFocus) {
				this.pendingFocus = false;
				this.parts.input.focus();
			}
		}
		private build() {
			const doc = this.ownerDocument;
			const p = config(this).testIdPrefix;
			const el = <K extends keyof HTMLElementTagNameMap>(tag: K, className = '') => {
				const node = doc.createElement(tag);
				node.className = className;
				return node;
			};
			const send = (intent: OfficeReadOnlyBannerIntent) =>
				emit(this, config(this).requestEvent, intent);
			const button = (testId: string, className: string, intent: OfficeReadOnlyBannerIntent) => {
				const node = el('button', className);
				node.type = 'button';
				node.dataset.testid = `${p}-${testId}`;
				node.addEventListener('click', () => send(intent));
				return node;
			};
			const banner = el('div', 'banner');
			banner.setAttribute('role', 'status');
			banner.setAttribute('part', 'banner');
			const title = el('strong');
			const message = el('span');
			const text = el('p', 'text');
			text.append(title, ': ', message);
			const editAnyway = button('edit-anyway', 'primary', { id: 'editAnyway' });
			const dismiss = button('dismiss', '', { id: 'dismiss' });
			const form = el('form', 'form');
			form.dataset.testid = `${p}-password-form`;
			const label = el('label', 'sr-only');
			const input = el('input');
			input.id = 'password';
			input.type = 'password';
			input.dataset.testid = `${p}-password-input`;
			label.htmlFor = input.id;
			const unlock = el('button', 'primary');
			unlock.type = 'submit';
			unlock.dataset.testid = `${p}-unlock`;
			const cancel = button('password-cancel', '', { id: 'cancelPassword' });
			const error = el('span', 'error');
			error.id = 'error';
			error.setAttribute('role', 'alert');
			error.dataset.testid = `${p}-password-error`;
			form.addEventListener('submit', (event) => {
				event.preventDefault();
				send({ id: 'submitPassword', password: input.value });
			});
			form.append(label, input, unlock, cancel, error);
			const lock = createIconSvg(doc);
			paintIcon(lock, 'lock');
			banner.append(lock, text, editAnyway, dismiss, form);
			return {
				banner,
				title,
				message,
				editAnyway,
				dismiss,
				form,
				label,
				input,
				unlock,
				cancel,
				error,
			};
		}
		private render(): void {
			const s = this.model;
			const l = s.labels ?? {};
			const p = this.parts;
			this.dataset.testid = `${config(this).testIdPrefix}-banner`;
			if (s.kind) this.dataset.kind = s.kind;
			else delete this.dataset.kind;
			p.title.textContent = l.title ?? 'Read-only recommended';
			p.message.textContent = s.message;
			p.editAnyway.textContent = l.editAnyway ?? 'Edit anyway';
			p.dismiss.textContent = l.dismiss ?? 'Dismiss';
			const open = s.passwordPromptOpen === true;
			p.editAnyway.hidden = p.dismiss.hidden = open;
			p.form.hidden = !open;
			p.label.textContent = l.passwordLabel ?? 'Password';
			p.input.placeholder = l.passwordPlaceholder ?? 'Password';
			p.unlock.textContent = l.unlock ?? 'Unlock';
			p.cancel.textContent = l.cancel ?? 'Cancel';
			p.input.disabled = p.unlock.disabled = s.checkingPassword === true;
			const failed = s.passwordError ?? null;
			p.input.setAttribute('aria-invalid', String(failed !== null));
			if (failed) p.input.setAttribute('aria-describedby', p.error.id);
			else p.input.removeAttribute('aria-describedby');
			p.error.hidden = failed === null;
			p.error.textContent = failed ?? '';
			if (open && !this.promptWasOpen) {
				if (this.isConnected) p.input.focus();
				else this.pendingFocus = true;
			} else if (!open && this.promptWasOpen) p.input.value = '';
			this.promptWasOpen = open;
		}
	}
	return OfficeUiReadOnlyBanner;
});

// ---------------------------------------------------------------------------------------------
// Paste Options
// ---------------------------------------------------------------------------------------------

export interface OfficePasteOption {
	/** Echoed back as `format` in the request event. */
	id: string;
	/** Translated label. */
	label: string;
}

export interface OfficePasteOptionsState {
	/** Right edge of the pasted object in viewport pixels. */
	left: number;
	/** Bottom edge of the pasted object in viewport pixels. */
	top: number;
	options: readonly OfficePasteOption[];
	/** Toolbar accessible name. */
	label?: string;
}

const PASTE_CSS = `
:host { position: fixed; z-index: 1100; display: block; }
:host([hidden]) { display: none !important; }
.toolbar { box-sizing: border-box; display: flex; align-items: center; gap: 2px; padding: 4px;
	border: 1px solid var(--office-border, #d1d5db); border-radius: 4px; background: var(--office-popover, var(--office-background, #fff));
	box-shadow: 0 10px 25px rgba(0, 0, 0, .2); font: 11px/1.2 var(--office-font, system-ui, sans-serif); }
button { box-sizing: border-box; min-height: 24px; padding: 4px 8px; border: 0; border-radius: 4px; background: transparent;
	color: var(--office-foreground, #1f2937); font: inherit; white-space: nowrap; cursor: pointer; touch-action: manipulation; }
button:hover { background: var(--office-surface, #f3f4f6); }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 1px; }
@media (pointer: coarse) { button { min-height: 44px; min-width: 44px; } }
@media (forced-colors: active) {
	.toolbar { border-color: CanvasText; background: Canvas; }
	button { color: ButtonText; } button:hover { background: Highlight; color: HighlightText; }
	button:focus-visible { outline-color: Highlight; }
}
`;

/**
 * `<office-ui-paste-options>`: the strip anchored 4px past a just-pasted object's bottom-right
 * corner (the host is the fixed box). A choice emits `office-paste-options-request`
 * `{ format }` then `office-paste-options-dismiss`; the first pointerdown or keydown outside
 * (Escape inside too) dismisses, armed one task after connecting so the paste gesture does not.
 */
export const definePasteOptions = definer('office-ui-paste-options', () => {
	class OfficeUiPasteOptions extends HTMLElement {
		static requestEvent = 'office-paste-options-request';
		static dismissEvent = 'office-paste-options-dismiss';
		static testIdPrefix = 'office-paste-options';
		private model: OfficePasteOptionsState = { left: 0, top: 0, options: [] };
		private readonly toolbar = this.ownerDocument.createElement('div');
		private disarm: (() => void) | undefined;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(PASTE_CSS));
			this.toolbar.className = 'toolbar';
			this.toolbar.setAttribute('role', 'toolbar');
			this.toolbar.tabIndex = -1;
			// Keep a press on the strip from reaching the canvas underneath.
			this.toolbar.addEventListener('mousedown', (event) => event.stopPropagation());
			root.append(this.toolbar);
		}
		get state(): OfficePasteOptionsState {
			return this.model;
		}
		set state(value: OfficePasteOptionsState) {
			this.model = value;
			this.render();
		}
		connectedCallback(): void {
			this.render();
			const view = this.ownerDocument.defaultView;
			if (!view) return;
			const dismiss = (event: Event) => {
				if (
					event.composedPath().includes(this.toolbar) &&
					(!(event instanceof KeyboardEvent) || event.key !== 'Escape')
				)
					return;
				emit(this, config(this).dismissEvent!);
			};
			const timer = view.setTimeout(() => {
				view.addEventListener('pointerdown', dismiss, true);
				view.addEventListener('keydown', dismiss, true);
			}, 0);
			this.disarm = () => {
				view.clearTimeout(timer);
				view.removeEventListener('pointerdown', dismiss, true);
				view.removeEventListener('keydown', dismiss, true);
			};
		}
		disconnectedCallback(): void {
			this.disarm?.();
			this.disarm = undefined;
		}
		private render(): void {
			const s = this.model;
			this.toolbar.setAttribute('aria-label', s.label ?? 'Paste Options');
			this.style.left = `${s.left + 4}px`;
			this.style.top = `${s.top + 4}px`;
			const doc = this.ownerDocument;
			const existing = [...this.toolbar.querySelectorAll<HTMLButtonElement>('button')];
			const same =
				existing.length === s.options.length &&
				existing.every((button, i) => button.dataset.format === s.options[i]!.id);
			if (!same)
				this.toolbar.replaceChildren(
					...s.options.map((option) => {
						const button = doc.createElement('button');
						button.type = 'button';
						button.dataset.format = option.id;
						button.addEventListener('click', () => {
							emit(this, config(this).requestEvent, { format: option.id });
							// A choice ends the follow-up, as in Office: ask the host to close it.
							emit(this, config(this).dismissEvent!);
						});
						return button;
					}),
				);
			this.toolbar.querySelectorAll<HTMLButtonElement>('button').forEach((button, i) => {
				const label = s.options[i]!.label;
				button.textContent = label;
				button.title = label;
			});
		}
	}
	return OfficeUiPasteOptions;
});
