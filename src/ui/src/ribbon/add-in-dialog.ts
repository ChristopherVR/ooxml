import type { OfficeAddInHost } from 'ooxml-core/automation/add-ins';
import { modalActiveElement } from '../dialog/focus';
import { showOfficeServiceDialog, type OfficeServiceText } from '../dialog/office-service-dialog';

export type AddInDialogMode = 'getAddIns' | 'myAddIns';
export type AddInDialogProduct = OfficeAddInHost | 'visio';

/** Trusted host commands for add-ins that are already implemented by the host. */
export interface OfficeAddInCatalogEntry {
	id: string;
	label: string;
	install: () => void | Promise<void>;
	remove?: () => void | Promise<void>;
}

/** An opt-in storefront offer. The dialog validates these bytes before asking for consent. */
export interface OfficeAddInOffer {
	id: string;
	label: string;
	read(signal: AbortSignal): File | Promise<File>;
}

export type OfficeAddInStorefront =
	| 'off'
	| { kind: 'microsoft'; url?: string; label?: string }
	| { kind: 'catalog'; entries: readonly OfficeAddInOffer[] };

export type AddInDialogHost = HTMLElement | ShadowRoot;

const CATALOG = Symbol.for('office-ui.add-in-catalog');
const STOREFRONT = Symbol.for('office-ui.add-in-storefront');
const OPENING = Symbol.for('office-ui.add-ins-opening');
const ENTRY_LIMIT = 128;
type HostSettings = AddInDialogHost & {
	[CATALOG]?: readonly OfficeAddInCatalogEntry[];
	[STOREFRONT]?: OfficeAddInStorefront;
	[OPENING]?: boolean;
};

/** Set or clear the trusted catalog for a single editor or host container. */
export function setOfficeAddInCatalog(
	host: AddInDialogHost,
	entries: readonly OfficeAddInCatalogEntry[],
): void {
	const target: HostSettings = host;
	if (entries.length) target[CATALOG] = entries.slice(0, ENTRY_LIMIT);
	else delete target[CATALOG];
}

/** Enable a storefront for one editor or host container. The default is off. */
export function setOfficeAddInStorefront(
	host: AddInDialogHost,
	storefront: OfficeAddInStorefront,
): void {
	const target: HostSettings = host;
	target[STOREFRONT] =
		storefront !== 'off' && storefront.kind === 'catalog'
			? { kind: 'catalog', entries: storefront.entries.slice(0, ENTRY_LIMIT) }
			: storefront;
}

export function addInSettings(host: AddInDialogHost) {
	let catalog: readonly OfficeAddInCatalogEntry[] | undefined;
	let storefront: OfficeAddInStorefront | undefined;
	for (let current: AddInDialogHost | null = host; current;) {
		const settings: HostSettings = current;
		catalog ??= settings[CATALOG];
		storefront ??= settings[STOREFRONT];
		const parent: Node | null = current instanceof ShadowRoot ? current.host : current.parentNode;
		current = parent instanceof HTMLElement || parent instanceof ShadowRoot ? parent : null;
	}
	return { catalog: catalog ?? [], storefront: storefront ?? 'off' };
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
	locale: string;
}

export async function openOfficeAddIns(
	host: AddInDialogHost,
	mode: AddInDialogMode,
	t: OfficeServiceText,
	product: AddInDialogProduct,
	locale: string,
): Promise<void> {
	const pending: HostSettings = host;
	const hasDialog = () =>
		[...host.children].some((child) => child.localName === 'office-ui-add-ins-dialog');
	if (hasDialog()) await Promise.resolve(); // Let a just-closed dialog detach.
	if (pending[OPENING] || hasDialog()) return;
	const opener = modalActiveElement(host.ownerDocument);
	const restore = addInReturnFocus(opener);
	pending[OPENING] = true;
	try {
		const { mountAddInsDialog } = await import('../dialog/add-ins-dialog');
		if (host.isConnected) mountAddInsDialog({ host, mode, t, product, locale }, restore);
	} catch {
		if (!host.isConnected) return;
		const dialog = showOfficeServiceDialog(host, t('Office Add-ins'), t('Close'));
		const status = host.ownerDocument.createElement('p');
		status.setAttribute('role', 'status');
		status.textContent = t('Could not load add-ins.');
		dialog.append(status);
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
