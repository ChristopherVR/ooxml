import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
import { defineCloudFiles } from '../controls';
import type { OfficeCloudFileOpenEvent, OfficeUiCloudFiles } from '../controls';
import { cloudFileLabels } from './cloud-labels';
import type { PageContext } from './backstage-pages';

/** The places File > Open lists: this device first, then one per cloud provider. */
export function renderOpenPlaces(
	{ handlers, t, content }: PageContext,
	providers: readonly CloudFileProvider[],
	device: () => HTMLElement[],
	heading: HTMLElement,
): void {
	defineCloudFiles();
	const nav = document.createElement('div');
	nav.className = 'dve-open-places';
	nav.setAttribute('role', 'group');
	nav.setAttribute('aria-label', t('cloud.providers'));
	const pane = document.createElement('div');
	pane.className = 'dve-open-pane';
	const places = [
		{ id: '', label: t('This device') },
		...providers.map((provider) => ({ id: provider.id, label: provider.label })),
	];
	const buttons = new Map<string, HTMLButtonElement>();
	const show = (id: string) => {
		for (const [key, button] of buttons) button.setAttribute('aria-pressed', String(key === id));
		const provider = providers.find((candidate) => candidate.id === id);
		if (!provider) return void pane.replaceChildren(...device());
		const files = document.createElement('office-ui-cloud-files') as OfficeUiCloudFiles;
		files.providers = [provider];
		files.accept = ['docx'];
		files.labels = cloudFileLabels(handlers.locale());
		files.locale = handlers.locale();
		files.addEventListener('office-cloud-file-open', (event) => {
			const { file } = (event as OfficeCloudFileOpenEvent).detail;
			void handlers.openCloudFile(file satisfies CloudOpenedFile);
		});
		pane.replaceChildren(files);
	};
	for (const place of places) {
		const button = document.createElement('button');
		button.type = 'button';
		button.className = 'dve-open-place';
		button.textContent = place.label;
		button.addEventListener('click', () => show(place.id));
		buttons.set(place.id, button);
		nav.append(button);
	}
	const layout = document.createElement('div');
	layout.className = 'dve-open-layout';
	layout.append(nav, pane);
	content.replaceChildren(heading, layout);
	show('');
}
