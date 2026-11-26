import type { OfficeAddInHost } from 'ooxml-core/automation/add-ins';
import {
	officeServiceLink,
	showOfficeServiceDialog,
	type OfficeServiceText,
} from '../dialog/office-service-dialog';

export type AddInDialogMode = 'getAddIns' | 'myAddIns';

/** A trusted host supplies the catalog and updates its ribbon tabs after changes. */
export interface OfficeAddInCatalogEntry {
	id: string;
	label: string;
	install: () => void | Promise<void>;
	remove?: () => void | Promise<void>;
}

// PowerPoint and the general UI entry are built separately and share the host node.
const ADD_IN_CATALOG = Symbol.for('office-ui.add-in-catalog');
type CatalogHost = (HTMLElement | ShadowRoot) & {
	[ADD_IN_CATALOG]?: readonly OfficeAddInCatalogEntry[];
};

/** Set or clear the locally available add-ins for a single editor or host container. */
export function setOfficeAddInCatalog(
	host: HTMLElement | ShadowRoot,
	entries: readonly OfficeAddInCatalogEntry[],
): void {
	const target: CatalogHost = host;
	if (entries.length) target[ADD_IN_CATALOG] = [...entries];
	else delete target[ADD_IN_CATALOG];
}

interface EditorAddIn {
	id: string;
	label: string;
	open: () => void;
}

function editorRibbons(scope: ParentNode): EditorAddIn[] {
	const ribbons = scope.querySelectorAll('office-ui-ribbon, pptx-ui-ribbon-tabs');
	return [...ribbons].flatMap((ribbon) => {
		const tabs = () => [
			...(ribbon.shadowRoot?.querySelectorAll<HTMLButtonElement>('button[role="tab"]') ?? []),
		];
		const installed = new Map<string, string>();
		for (const panel of [...ribbon.children] as HTMLElement[]) {
			const { ribbonTab, label } = panel.dataset;
			if (ribbonTab && label && panel.hasAttribute('data-add-in')) installed.set(ribbonTab, label);
		}
		for (const tab of tabs()) {
			const id = tab.dataset.ribbonAddInTab;
			const label = tab.textContent?.trim();
			if (id && label) installed.set(id, label);
		}
		return [...installed].map(([id, label]): EditorAddIn => ({
			id,
			label,
			open: () => {
				const tab = tabs().find((button) => button.dataset.tab === id);
				if (tab?.isConnected) {
					tab.click();
					tab.focus();
				}
			},
		}));
	});
}

/** Find the containing editor's ribbon even when the Insert control is in an overlay. */
function editorAddIns(host: HTMLElement | ShadowRoot): EditorAddIn[] {
	if (host.nodeType === 11) return editorRibbons(host);
	for (let scope: HTMLElement | null = host as HTMLElement; scope; scope = scope.parentElement) {
		if (scope === host.ownerDocument.body) break;
		const tabs = editorRibbons(scope);
		if (tabs.length || scope.querySelector('office-ui-ribbon, pptx-ui-ribbon-tabs')) return tabs;
	}
	const root = host.getRootNode();
	return root.nodeType === 11 ? editorRibbons(root as ShadowRoot) : [];
}

function addInCatalog(host: HTMLElement | ShadowRoot): readonly OfficeAddInCatalogEntry[] {
	for (let scope: HTMLElement | ShadowRoot | null = host; scope;) {
		const entries = (scope as CatalogHost)[ADD_IN_CATALOG];
		if (entries) return entries;
		if (scope instanceof ShadowRoot) {
			scope = scope.host instanceof HTMLElement ? scope.host : null;
		} else {
			const parent: ParentNode | null = scope.parentNode;
			scope = parent instanceof HTMLElement || parent instanceof ShadowRoot ? parent : null;
		}
	}
	return [];
}

export function openOfficeAddIns(
	host: HTMLElement | ShadowRoot,
	mode: AddInDialogMode,
	t: OfficeServiceText,
	product: OfficeAddInHost,
	locale: string,
): void {
	const doc = host.ownerDocument;
	const tabs = editorAddIns(host);
	const catalog = addInCatalog(host);
	const dialog = showOfficeServiceDialog(
		host,
		t(mode === 'getAddIns' ? 'Get Add-ins' : 'My Add-ins'),
		t('Close'),
	);
	const list = doc.createElement('ul');
	const status = doc.createElement('p');
	status.setAttribute('role', 'status');
	status.hidden = true;
	let updating = false;
	const changeButton = (label: string, name: string, change: () => void | Promise<void>) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.textContent = t(label);
		button.setAttribute('aria-label', `${t(label)} ${name}`);
		button.addEventListener('click', async () => {
			if (updating) return;
			updating = true;
			button.disabled = true;
			status.hidden = true;
			try {
				await change();
				dialog.close();
			} catch {
				status.textContent = t('Could not update the add-in.');
				status.hidden = false;
				button.disabled = false;
				button.focus();
			} finally {
				updating = false;
			}
		});
		return button;
	};
	let empty: HTMLElement | undefined;
	let marketplace: HTMLElement | undefined;
	if (mode === 'getAddIns') {
		const installed = new Set(tabs.map(({ id }) => id));
		for (const entry of catalog) {
			if (installed.has(entry.id)) continue;
			const item = doc.createElement('li');
			item.append(`${entry.label} `, changeButton('Install', entry.label, entry.install));
			list.append(item);
		}
		if (list.childElementCount) dialog.append(list, status);
		const note = doc.createElement('p');
		note.textContent = t(
			'Browse add-ins for Microsoft Office. This editor runs add-ins supplied by its host.',
		);
		const url = 'https://marketplace.microsoft.com/en-us/marketplace/apps?product=office';
		marketplace = doc.createElement('div');
		marketplace.append(note, officeServiceLink(doc, t('Browse Microsoft Marketplace'), url));
		dialog.append(marketplace);
	} else if (tabs.length) {
		for (const tab of tabs) {
			const item = doc.createElement('li');
			const button = doc.createElement('button');
			button.type = 'button';
			button.textContent = tab.label;
			button.addEventListener('click', () => {
				if (updating) return;
				dialog.close();
				tab.open();
			});
			item.append(button);
			const remove = catalog.find(({ id }) => id === tab.id)?.remove;
			if (remove) item.append(' ', changeButton('Remove', tab.label, remove));
			list.append(item);
		}
		dialog.append(list, status);
	} else {
		empty = doc.createElement('p');
		empty.textContent = t('No add-ins are available in this editor.');
		dialog.append(empty);
	}
	dialog.show();
	void import('./add-in-local-dialog').then(
		({ mountLocalAddIns }) => {
			if (dialog.open && dialog.isConnected)
				mountLocalAddIns(dialog, host, mode, product, locale, t, { empty, marketplace });
		},
		() => {
			if (!dialog.open || !dialog.isConnected) return;
			status.textContent = t('Could not update the add-in.');
			status.hidden = false;
			dialog.append(status);
		},
	);
}
