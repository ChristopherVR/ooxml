/**
 * File > Open with cloud places: the Open page renders the shared places element with the
 * providers and the translated labels, and hands a file opened there to the viewer.
 */
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../../i18n';
import type { BackstageCloudFiles, RibbonFileHandlers } from '../ribbon-types';
import { createFileTab } from './file-tab';

const t = createTranslator('en');

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

function handlers(cloud?: BackstageCloudFiles): RibbonFileHandlers {
	const bag = {
		getRecentPresentationsCount: () => 0,
		...(cloud ? { getCloudFiles: () => cloud } : {}),
	};
	return new Proxy(bag as unknown as RibbonFileHandlers, {
		get(target, prop) {
			if (prop in target) {
				return (target as unknown as Record<string | symbol, unknown>)[prop];
			}
			return typeof prop === 'string' && prop.startsWith('get') ? undefined : vi.fn();
		},
	});
}

function openPage(el: HTMLElement): void {
	const nav = [...el.querySelectorAll<HTMLButtonElement>('[data-pptx-backstage-nav-item]')].find(
		(item) => item.textContent === t('pptx.backstage.nav.open'),
	);
	nav?.click();
}

afterEach(() => {
	document.body.replaceChildren();
});

interface PlacesElement extends HTMLElement {
	providers: unknown[];
	accept: string[];
	labels: Record<string, string>;
	locale: string | undefined;
}

describe('File > Open cloud places', () => {
	it('keeps the Open page as it was without providers', () => {
		const tab = createFileTab(document, t, handlers(), vi.fn());
		document.body.appendChild(tab.el);
		openPage(tab.el);
		const places = tab.el.querySelector<PlacesElement>('pptx-ui-open-places')!;
		expect(places.providers).toEqual([]);
		expect(tab.el.textContent).toContain(t('pptx.backstage.browseDevice'));
	});

	it('hands the providers to the places element and opens a provider file', () => {
		const open = vi.fn();
		const onClose = vi.fn();
		const tab = createFileTab(
			document,
			t,
			handlers({ providers: [provider], locale: 'en', open }),
			onClose,
		);
		document.body.appendChild(tab.el);
		openPage(tab.el);
		const places = tab.el.querySelector<PlacesElement>('pptx-ui-open-places')!;
		expect(places.providers).toEqual([provider]);
		expect(places.accept).toEqual(['pptx']);
		expect(places.labels.thisDevice).toBe(t('pptx.backstage.thisDevice'));
		expect(places.labels.places).toBe(t('pptx.backstage.places'));
		expect(places.locale).toBe('en');
		// The device content is the element's light DOM, so the element can show or hide it.
		expect(places.textContent).toContain(t('pptx.backstage.browseDevice'));
		const file = { fileName: 'Cloud.pptx' } as CloudOpenedFile;
		places.dispatchEvent(
			new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
		);
		expect(open).toHaveBeenCalledWith(file);
		expect(onClose).toHaveBeenCalled();
	});

	it('offers the providers on the Open page only', () => {
		const tab = createFileTab(
			document,
			t,
			handlers({ providers: [provider], locale: 'en', open: vi.fn() }),
			vi.fn(),
		);
		document.body.appendChild(tab.el);
		expect(tab.el.querySelector<PlacesElement>('pptx-ui-open-places')?.providers).toEqual([]);
		openPage(tab.el);
		expect(tab.el.querySelector<PlacesElement>('pptx-ui-open-places')?.providers).toEqual([
			provider,
		]);
	});
});
