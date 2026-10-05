import type { EditorContext } from 'ooxml-core/xlsx/ui';

/**
 * The dialog host: the element in the shadow root dialogs append their `office-ui-dialog` to, so
 * dialog code does not need to know the shell layout.
 */
export function dialogHost(ctx: EditorContext): HTMLElement {
	const existing = ctx.root.querySelector<HTMLElement>('[data-dialog-host]');
	if (existing) return existing;
	const host = ctx.host.ownerDocument.createElement('div');
	host.dataset.dialogHost = '';
	ctx.root.append(host);
	return host;
}
