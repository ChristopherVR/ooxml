import type { OfficeAddInHost } from 'ooxml-core/automation/add-ins';
import { normalizeLocale } from 'ooxml-core/i18n';
import { composedWithin } from '../composed';
import { modalActiveElement } from '../dialog/focus';
import { showOfficeServiceDialog, type OfficeServiceText } from '../dialog/office-service-dialog';

export type AddInDialogMode = 'getAddIns' | 'myAddIns';
export type AddInDialogProduct = OfficeAddInHost | 'visio';

export type AddInDialogHost = HTMLElement | ShadowRoot;

const OPENING = Symbol.for('office-ui.add-ins-opening');
const LOAD_ERRORS = new Map([
	['fr', 'Impossible de charger les compléments.'],
	['de', 'Add-Ins konnten nicht geladen werden.'],
	['es', 'No se pudieron cargar los complementos.'],
	['zh-CN', '无法加载加载项。'],
	['ja', 'アドインを読み込めませんでした。'],
	['ko', '추가 기능을 불러오지 못했습니다.'],
]);
const LOCALES = [...LOAD_ERRORS.keys()];
type PendingHost = AddInDialogHost & {
	[OPENING]?: boolean;
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
}

export async function openOfficeAddIns(
	host: AddInDialogHost,
	mode: AddInDialogMode,
	t: OfficeServiceText,
	product: AddInDialogProduct,
	locale: string,
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
			mountAddInsDialog({ host, mode, t, product, manifestLocale: locale, scopes }, restore);
		}
	} catch {
		if (!host.isConnected) return;
		const dialog = showOfficeServiceDialog(host, t('Office Add-ins'), t('Close'));
		dialog.setAttribute('data-office-add-ins-fallback', '');
		const status = host.ownerDocument.createElement('p');
		status.setAttribute('role', 'status');
		const language = normalizeLocale(locale, LOCALES, 'en');
		status.textContent = LOAD_ERRORS.get(language) ?? 'Could not load add-ins.';
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
