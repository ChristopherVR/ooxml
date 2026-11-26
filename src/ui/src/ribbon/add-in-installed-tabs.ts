import { MANIFEST_TAB_PREFIX } from './add-in-manifest-tabs';

export interface EditorAddIn {
	id: string;
	label: string;
	open: () => void;
}

export function editorAddIns(scope: ParentNode): EditorAddIn[] {
	const ribbons = scope.querySelectorAll('office-ui-ribbon, pptx-ui-ribbon-tabs');
	return [...ribbons].flatMap((ribbon) => {
		const tabs = () => [
			...(ribbon.shadowRoot?.querySelectorAll<HTMLButtonElement>('button[role="tab"]') ?? []),
		];
		const installed = new Map<string, string>();
		for (const panel of ribbon.querySelectorAll<HTMLElement>(':scope > [data-add-in]')) {
			const { ribbonTab, label } = panel.dataset;
			if (ribbonTab && label && !ribbonTab.startsWith(MANIFEST_TAB_PREFIX))
				installed.set(ribbonTab, label);
		}
		for (const tab of tabs()) {
			const id = tab.dataset.ribbonAddInTab;
			const label = tab.textContent?.trim();
			if (id && label && !id.startsWith(MANIFEST_TAB_PREFIX)) installed.set(id, label);
		}
		return [...installed].map(([id, label]) => ({
			id,
			label,
			open: () => {
				const tab = tabs().find((button) => button.dataset.tab === id);
				if (tab?.isConnected) {
					tab.click();
					tab.focus({ preventScroll: true });
				}
			},
		}));
	});
}
