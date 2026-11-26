import type { LocalAddIn } from 'ooxml-core/automation/add-ins';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import { modalActiveElement } from '../dialog/focus';
import type { OpenAddInsOptions } from './add-in-dialog';
import { addInPreferences } from './add-in-preferences';
import { addInRuntime, allowedPaneUrl } from './add-in-runtime-settings';

export function addInLauncher(
	options: OpenAddInsOptions,
	installed: ReadonlyMap<string, LocalAddIn>,
	t: OfficeServiceText,
) {
	const { host, product, scopes, selection } = options;
	const store = addInPreferences(scopes);
	function pane(addIn: LocalAddIn, source: string | null = addIn.manifest.pane) {
		const runtime = addInRuntime(scopes);
		if (addIn.manifest.identity)
			return { reason: t('Microsoft sign-in and access tokens are unavailable.') };
		if (!runtime)
			return { reason: t('Installed for this editor. Running add-ins is unavailable.') };
		if (runtime.frameSrcEnforced !== true)
			return { reason: t('The host must enforce a frame policy to run add-ins.') };
		if (!source) return { reason: t('This add-in has no compatible task pane.') };
		const url = allowedPaneUrl(source, runtime);
		if (!url || url.origin === host.ownerDocument.location?.origin)
			return { reason: t('The host has not approved this task pane origin.') };
		return { url, runtime };
	}
	async function open(addIn: LocalAddIn, source?: string, restoreFocus?: () => void) {
		if (product === 'visio') return;
		const { manifest } = addIn;
		const opener = modalActiveElement(host.ownerDocument);
		const live = pane(addIn, source);
		if (!live.url) return;
		const access = (live.runtime.selection ?? selection)?.();
		const { openAddInPane } = await import('./add-in-pane');
		const active = () => {
			const current = pane(addIn, source);
			return (
				host.isConnected &&
				installed.get(manifest.id) === addIn &&
				addInPreferences(scopes) === store &&
				current.url?.href === live.url.href &&
				current.runtime === live.runtime
			);
		};
		if (!active()) return;
		openAddInPane(host, {
			id: manifest.id,
			url: live.url,
			name: manifest.name,
			product,
			permissions: manifest.permissions,
			selection: access,
			active,
			scopes,
			t,
			restoreFocus:
				restoreFocus ??
				(() => {
					if (opener instanceof HTMLElement && opener.isConnected) opener.focus();
				}),
		});
	}
	const close = (id: string) => {
		for (const child of host.children) {
			if (
				child instanceof HTMLElement &&
				child.dataset.addInId === id &&
				child.localName === 'office-ui-add-in-pane'
			)
				child.remove();
		}
	};
	return { pane, open, close };
}
