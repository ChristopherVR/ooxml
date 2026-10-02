// Dialog registry: dialogs are opened by name so commands, the grid and the shell share them
// without importing one another. Implementations live in `dialogs/` (UI-COMMANDS).
import type { EditorContext } from './context';

export type DialogOpener = (ctx: EditorContext, props?: unknown) => Promise<unknown>;

export interface DialogRegistry {
	register(name: string, open: DialogOpener): void;
	/** Resolves undefined when cancelled or when no dialog has that name. */
	open<T = unknown>(name: string, props?: unknown): Promise<T | undefined>;
}

export function createDialogRegistry(
	context: () => EditorContext,
	onError?: (name: string, error: unknown) => void,
): DialogRegistry {
	const openers = new Map<string, DialogOpener>();
	return {
		register(name, open) {
			openers.set(name, open);
		},
		async open<T = unknown>(name: string, props?: unknown): Promise<T | undefined> {
			const opener = openers.get(name);
			if (!opener) return undefined;
			try {
				const result = await opener(context(), props);
				return (result ?? undefined) as T | undefined;
			} catch (error) {
				onError?.(name, error);
				return undefined;
			}
		},
	};
}

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
