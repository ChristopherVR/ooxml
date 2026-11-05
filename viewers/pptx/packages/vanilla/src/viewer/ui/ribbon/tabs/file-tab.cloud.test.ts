/**
 * File > Open with cloud places: the Open page lists "This device" and one place per provider,
 * shows the provider's file browser, and hands a file opened there to the viewer.
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

describe('File > Open cloud places', () => {
	it('keeps the Open page as it was without providers', () => {
		const tab = createFileTab(document, t, handlers(), vi.fn());
		document.body.appendChild(tab.el);
		openPage(tab.el);
		expect(tab.el.querySelector('[data-pptx-open-place]')).toBeNull();
		expect(tab.el.textContent).toContain(t('pptx.backstage.browseDevice'));
	});

	it('lists this device and the provider, and opens a provider file', () => {
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
		const tablist = tab.el.querySelector('[role="tablist"]');
		expect(tablist?.getAttribute('aria-label')).toBe(t('pptx.backstage.places'));
		const places = [...tab.el.querySelectorAll<HTMLButtonElement>('[data-pptx-open-place]')];
		expect(places.map((place) => place.textContent)).toStrictEqual([
			t('pptx.backstage.thisDevice'),
			'OneDrive',
		]);
		expect(places[0]?.getAttribute('aria-selected')).toBe('true');
		expect(tab.el.querySelector('pptx-ui-cloud-files')).toBeNull();
		places[1]?.click();
		const place = tab.el.querySelector('[data-pptx-cloud-place="onedrive"]');
		expect(place?.querySelector('pptx-ui-cloud-files')).not.toBeNull();
		expect(tab.el.textContent).not.toContain(t('pptx.backstage.browseDevice'));
		expect(
			tab.el
				.querySelector('[data-pptx-open-place="cloud:onedrive"]')
				?.getAttribute('aria-selected'),
		).toBe('true');
		const file = { fileName: 'Cloud.pptx' } as CloudOpenedFile;
		place
			?.querySelector('pptx-ui-cloud-files')
			?.dispatchEvent(
				new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
			);
		expect(open).toHaveBeenCalledWith(file);
		expect(onClose).toHaveBeenCalled();
	});

	it('returns to the device search and recent list from a cloud place', () => {
		const tab = createFileTab(
			document,
			t,
			handlers({ providers: [provider], locale: 'en', open: vi.fn() }),
			vi.fn(),
		);
		document.body.appendChild(tab.el);
		openPage(tab.el);
		tab.el.querySelector<HTMLButtonElement>('[data-pptx-open-place="cloud:onedrive"]')?.click();
		tab.el.querySelector<HTMLButtonElement>('[data-pptx-open-place="device"]')?.click();
		expect(tab.el.querySelector('pptx-ui-cloud-files')).toBeNull();
		expect(tab.el.textContent).toContain(t('pptx.backstage.browseDevice'));
	});
});
