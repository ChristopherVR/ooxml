import { createDialogSession } from '../../dialog/dialog-session';
// The modal frame every UI-COMMANDS dialog uses: an `office-ui-dialog` (focus trap, Escape,
// return focus) with OK / Cancel in its footer, Enter to confirm from a field, and a promise that
// resolves with the dialog's result or undefined when it is cancelled.
import { defineDialog } from '../../index';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { dialogHost } from './host';
import { DIALOG_CSS } from './styles';

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
	const opener = ctx.root.activeElement ?? doc.activeElement;
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

	const session = createDialogSession<T>({
		element: dialog,
		confirm: ok,
		cancel,
		...(spec.submit ? { submit: spec.submit } : {}),
		finish: () => {
			dialog.removeAttribute('open');
			dialog.remove();
			if (ctx.grid()) ctx.grid()?.focus();
			else if (opener instanceof doc.defaultView!.HTMLElement && opener.isConnected)
				opener.focus({ preventScroll: true });
		},
		error: (error) =>
			ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'error'),
	});
	dialogHost(ctx).append(dialog);
	dialog.setAttribute('open', '');
	handle?.({ element: dialog, close: () => session.close() });
	spec.opened?.();
	return session.result;
}
