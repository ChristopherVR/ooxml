import type { OfficeAddInHost } from 'ooxml-core/automation/add-ins';
import { composedWithin } from '../composed';
import { modalActiveElement } from '../dialog/focus';
import { showOfficeServiceDialog, type OfficeServiceText } from '../dialog/office-service-dialog';
import { localAddIns } from './add-in-installations';
import { ADD_IN_PREFERENCES_CHANGED, addInPreferences } from './add-in-preferences';
import { ADD_IN_RUNTIME_CHANGED, type OfficeAddInSelection } from './add-in-runtime-settings';

export type AddInDialogMode = 'getAddIns' | 'myAddIns';
export type AddInDialogProduct = OfficeAddInHost | 'visio';

export type AddInDialogHost = HTMLElement | ShadowRoot;

const RESTORE = Symbol.for('office-ui.mounted-add-in-restore');
const OPENING = Symbol.for('office-ui.add-ins-opening');
type PendingHost = AddInDialogHost & {
	[OPENING]?: boolean;
	[RESTORE]?: Promise<void>;
};

function addInScopes(host: AddInDialogHost, opener: Element | null): AddInDialogHost[] {
	const scopes: AddInDialogHost[] = [];
	const start = opener instanceof HTMLElement && composedWithin(opener, host) ? opener : host;
	for (let current: AddInDialogHost | null = start; current;) {
		scopes.push(current);
		const parent: Node | null = current instanceof ShadowRoot ? current.host : current.parentNode;
		current = parent instanceof HTMLElement || parent instanceof ShadowRoot ? parent : null;
	}
	return scopes;
}

/** PowerPoint may render ribbon controls inside a shadow tree or a standalone toolbar. */
export function officeAddInsMount(control: HTMLElement): AddInDialogHost | null {
	const editor = control.closest<HTMLElement>('[data-pptx-editor-chrome], [data-pptx-viewer]');
	if (editor) return editor;
	const root = control.getRootNode();
	return root instanceof ShadowRoot ? root : control.parentElement;
}

/** Register the on-demand add-in dialog in a browser or custom element registry. */
export async function defineOfficeAddInsDialog(
	registry = globalThis.customElements,
): Promise<void> {
	if (!registry) return;
	const { defineAddInsDialog } = await import('../dialog/add-ins-dialog');
	defineAddInsDialog(registry);
}

export interface OpenAddInsOptions {
	host: AddInDialogHost;
	mode: AddInDialogMode;
	t: OfficeServiceText;
	product: AddInDialogProduct;
	manifestLocale: string;
	scopes: readonly AddInDialogHost[];
	selection?: () => OfficeAddInSelection | undefined | Promise<OfficeAddInSelection | undefined>;
}

export function mountedAddInRestore(host: PendingHost) {
	return host[RESTORE];
}

export function watchOfficeAddIns(
	host: PendingHost,
	t: OfficeServiceText,
	product: OfficeAddInHost,
	locale: () => string,
	selection?: OpenAddInsOptions['selection'],
): () => void {
	const scopes = addInScopes(host, host.querySelector('office-ui-ribbon, pptx-ui-ribbon-tabs'));
	const listeners = new AbortController();
	let pending: AbortController | undefined;
	let activeStore = addInPreferences(scopes);
	let runtimeRefresh: AbortController | undefined;

	const options = (): OpenAddInsOptions => ({
		host,
		mode: 'myAddIns',
		t,
		product,
		manifestLocale: locale(),
		scopes,
		...(selection && { selection }),
	});

	const restore = () => {
		pending?.abort();
		delete host[RESTORE];
		const store = addInPreferences(scopes);
		const changed = activeStore !== store;
		if (changed) {
			activeStore = store;
			localAddIns(host).clear();
			host.querySelector('office-ui-add-in-pane')?.remove();
			host.querySelector('office-ui-add-ins-dialog')?.remove();
		}
		if (!store && !changed && !localAddIns(host).size) return;
		runtimeRefresh?.abort();
		const { signal } = (pending = new AbortController());
		host[RESTORE] = (async () => {
			const files = store ? await store.read(product) : [];
			if (signal.aborted || (!changed && !files.length && !localAddIns(host).size)) return;
			const { restoreMountedAddIns } = await import('../dialog/add-ins-session');
			await restoreMountedAddIns(options(), signal, files, changed);
		})();
		// The dialog can report a preference read error if the user opens it.
		void host[RESTORE].catch(() => {});
	};
	const refreshRuntime = () => {
		runtimeRefresh?.abort();
		if (!localAddIns(host).size && !host[RESTORE]) return;
		const { signal } = (runtimeRefresh = new AbortController());
		void (async () => {
			await host[RESTORE];
			if (signal.aborted) return;
			const { refreshMountedAddIns } = await import('../dialog/add-ins-session');
			await refreshMountedAddIns(options(), signal);
		})().catch(() => {});
	};
	for (const scope of scopes) {
		scope.addEventListener(ADD_IN_PREFERENCES_CHANGED, restore, { signal: listeners.signal });
		scope.addEventListener(ADD_IN_RUNTIME_CHANGED, refreshRuntime, { signal: listeners.signal });
	}
	restore();
	return () => {
		listeners.abort();
		runtimeRefresh?.abort();
		pending?.abort();
		delete host[RESTORE];
	};
}

export async function openOfficeAddIns(
	host: AddInDialogHost,
	mode: AddInDialogMode,
	t: OfficeServiceText,
	product: AddInDialogProduct,
	locale: string,
	selection?: OpenAddInsOptions['selection'],
): Promise<void> {
	const pending: PendingHost = host;
	if (
		pending[OPENING] ||
		[...host.children].some((child) =>
			child.matches('office-ui-add-ins-dialog, [data-office-add-ins-fallback]'),
		)
	)
		return;
	const opener = modalActiveElement(host.ownerDocument);
	// The command may leave a folded popup before the dialog module finishes loading.
	const scopes = addInScopes(host, opener);
	const restore = addInReturnFocus(opener);
	pending[OPENING] = true;
	try {
		const { mountAddInsDialog } = await import('../dialog/add-ins-dialog');
		if (host.isConnected) {
			mountAddInsDialog(
				{ host, mode, t, product, manifestLocale: locale, scopes, ...(selection && { selection }) },
				restore,
			);
		}
	} catch {
		const fallback = await import('./add-in-load-error').catch(() => undefined);
		if (!host.isConnected) return;
		const dialog = showOfficeServiceDialog(host, t('Office Add-ins'), t('Close'));
		dialog.setAttribute('data-office-add-ins-fallback', '');
		const status = host.ownerDocument.createElement('p');
		status.setAttribute('role', 'status');
		status.textContent = fallback?.addInLoadError(locale) ?? t('Could not load add-ins.');
		dialog.append(status);
		dialog.addEventListener('office-dialog-closed', (event) => {
			if (event.target === dialog) restore();
		});
		dialog.show();
	} finally {
		delete pending[OPENING];
	}
}

function addInReturnFocus(opener: Element | null): () => void {
	const root = opener?.getRootNode();
	const control =
		root instanceof ShadowRoot && root.host instanceof HTMLElement ? root.host : opener;
	const group = control?.closest('[data-ribbon-group], .ribbon-group');
	const pop = control?.closest('.ribbon-overflow-panel');
	const collapsed = pop?.parentElement?.querySelector<HTMLButtonElement>(
		'.ribbon-overflow-button[aria-expanded="true"]',
	);
	return () => {
		if (opener instanceof HTMLElement && opener.isConnected) {
			opener.focus({ preventScroll: true });
			if (modalActiveElement(opener.ownerDocument) === opener) return;
		}
		if (collapsed?.isConnected && collapsed.getClientRects().length) {
			collapsed.focus({ preventScroll: true });
			return;
		}
		const face =
			group?.querySelector<HTMLElement>(
				':scope > .ribbon-overflow-button, :scope > [data-pptx-chrome="ribbon-collapse"]',
			) ?? group?.shadowRoot?.querySelector<HTMLElement>('button.face');
		face?.focus({ preventScroll: true });
	};
}
