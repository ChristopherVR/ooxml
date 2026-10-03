// The modal frame every UI-COMMANDS dialog uses: an `office-ui-dialog` (focus trap, Escape,
// return focus) with OK / Cancel in its footer, Enter to confirm from a field, and a promise that
// resolves with the dialog's result or undefined when it is cancelled.
import { defineDialog } from 'ooxml-ui';
import type { EditorContext } from '../context.js';
import { dialogHost } from '../dialogs.js';
import { DIALOG_CSS } from './styles.js';

export interface DialogSpec<T> {
	/** English heading (translated here). */
	heading: string;
	body: HTMLElement | HTMLElement[];
	/** English label of the confirm button; default `OK`. Null hides OK (a Close-only dialog). */
	okLabel?: string | null;
	cancelLabel?: string;
	/** Extra footer buttons, before Cancel / OK. */
	buttons?: HTMLButtonElement[];
	/** Validates and applies; return undefined to keep the dialog open. */
	submit?(): T | undefined | Promise<T | undefined>;
	/** Called once the dialog is in the document and open (set the initial focus here). */
	opened?(): void;
	/** Wider frame for tabbed dialogs. */
	wide?: boolean;
	/** `data-dialog` name, for tests and styling. */
	name: string;
}

export interface OpenDialog {
	element: HTMLElement;
	close(): void;
}

const STYLE_ID = 'xve-command-dialog-css';

function ensureStyles(ctx: EditorContext): void {
	if (ctx.root.getElementById?.(STYLE_ID) ?? ctx.root.querySelector(`#${STYLE_ID}`)) return;
	const style = ctx.host.ownerDocument.createElement('style');
	style.id = STYLE_ID;
	style.textContent = DIALOG_CSS;
	ctx.root.append(style);
}

/** A plain footer / body button (`primary` gets the accent look). */
export function button(ctx: EditorContext, label: string, primary = false): HTMLButtonElement {
	const el = ctx.host.ownerDocument.createElement('button');
	el.type = 'button';
	el.textContent = ctx.t(label);
	el.className = primary ? 'xve-btn xve-btn-primary' : 'xve-btn';
	return el;
}

/** Opens a dialog and resolves with its result (undefined when cancelled or closed). */
export function showDialog<T>(
	ctx: EditorContext,
	spec: DialogSpec<T>,
	handle?: (open: OpenDialog) => void,
): Promise<T | undefined> {
	defineDialog(ctx.host.ownerDocument.defaultView?.customElements);
	ensureStyles(ctx);
	const doc = ctx.host.ownerDocument;
	const dialog = doc.createElement('office-ui-dialog');
	dialog.setAttribute('heading', ctx.t(spec.heading));
	dialog.dataset.dialog = spec.name;
	dialog.className = spec.wide ? 'xve-cmd-dialog xve-cmd-dialog-wide' : 'xve-cmd-dialog';
	const body = doc.createElement('div');
	body.className = 'xve-dialog-body';
	body.append(...(Array.isArray(spec.body) ? spec.body : [spec.body]));
	const footer = doc.createElement('div');
	footer.slot = 'footer';
	footer.className = 'xve-dialog-footer';
	const ok = spec.okLabel === null ? undefined : button(ctx, spec.okLabel ?? 'OK', true);
	const cancel = button(ctx, spec.cancelLabel ?? (ok ? 'Cancel' : 'Close'));
	footer.append(...(spec.buttons ?? []), cancel, ...(ok ? [ok] : []));
	dialog.append(body, footer);

	return new Promise<T | undefined>((resolve) => {
		let done = false;
		const finish = (result: T | undefined): void => {
			if (done) return;
			done = true;
			dialog.removeAttribute('open');
			dialog.remove();
			resolve(result);
		};
		// One submit at a time: a double Enter or double click must not apply an async submit twice.
		// A synchronous submit settles at once; an async one disables OK until it settles.
		let submitting = false;
		const failed = (error: unknown) =>
			ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'error');
		const settle = (result: T | undefined) => {
			if (result !== undefined) finish(result);
		};
		const confirm = async (): Promise<void> => {
			if (done || submitting) return;
			if (!spec.submit) return finish(undefined);
			let outcome: T | undefined | Promise<T | undefined>;
			try {
				outcome = spec.submit();
			} catch (error) {
				return failed(error);
			}
			if (!(outcome instanceof Promise)) return settle(outcome);
			submitting = true;
			if (ok) ok.disabled = true;
			try {
				settle(await outcome);
			} catch (error) {
				failed(error);
			} finally {
				submitting = false;
				if (ok && !done) ok.disabled = false;
			}
		};
		dialog.addEventListener('office-dialog-close', () => finish(undefined));
		cancel.addEventListener('click', () => finish(undefined));
		ok?.addEventListener('click', () => void confirm());
		dialog.addEventListener('keydown', (event) => {
			const el = event.target as HTMLElement | null;
			if (event.key === 'Escape' && !event.defaultPrevented) {
				// office-ui-dialog handles Escape when it is upgraded; this covers the rest.
				event.preventDefault();
				finish(undefined);
			} else if (
				event.key === 'Enter' &&
				ok &&
				el instanceof doc.defaultView!.HTMLInputElement &&
				el.type !== 'checkbox' &&
				el.type !== 'radio' &&
				el.type !== 'file'
			) {
				event.preventDefault();
				void confirm();
			}
		});
		dialogHost(ctx).append(dialog);
		dialog.setAttribute('open', '');
		handle?.({ element: dialog, close: () => finish(undefined) });
		spec.opened?.();
	});
}
