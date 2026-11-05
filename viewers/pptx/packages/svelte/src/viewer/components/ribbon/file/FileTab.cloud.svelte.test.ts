/**
 * File > Open with cloud places: the Open page lists "This device" and one place per provider,
 * shows the provider's file browser, and hands a file opened there to the viewer.
 */
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { ViewerCloudFiles } from '../../../state/viewer-cloud-files-context';
import FileTabCloudHarness from './FileTabCloudHarness.svelte';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

function render(
	cloud: ViewerCloudFiles | undefined,
	overrides: { onclose?: () => void; oncreatepresentation?: (id: string) => void } = {},
): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const noop = vi.fn();
	const instance = mount(FileTabCloudHarness, {
		target,
		props: {
			cloud,
			onclose: overrides.onclose ?? noop,
			oncreatepresentation: overrides.oncreatepresentation ?? noop,
		},
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

function openPage(target: HTMLElement): void {
	const open = [
		...target.querySelectorAll<HTMLButtonElement>('[data-pptx-backstage-nav-item]'),
	].find((button) => button.textContent?.trim() === 'Open');
	open!.click();
	flushSync();
}

describe('FileTab open places', () => {
	it('keeps the Open page as it was without providers', () => {
		const target = render(undefined);
		openPage(target);
		expect(target.querySelector('[data-pptx-open-place]')).toBeNull();
		expect(target.textContent).toContain('Browse');
	});

	it('lists this device and the provider, and opens a provider file', () => {
		const open = vi.fn();
		const onclose = vi.fn();
		const target = render({ providers: [provider], open, clear: vi.fn() }, { onclose });
		openPage(target);
		const tablist = target.querySelector('[role="tablist"]');
		expect(tablist?.getAttribute('aria-label')).toBe('Places');
		const places = [...target.querySelectorAll<HTMLButtonElement>('[data-pptx-open-place]')];
		expect(places.map((place) => place.textContent?.trim())).toEqual(['This device', 'OneDrive']);
		expect(places.map((place) => place.dataset.pptxOpenPlace)).toEqual([
			'device',
			'cloud:onedrive',
		]);
		expect(target.querySelector('pptx-ui-cloud-files')).toBeNull();
		places[1]!.click();
		flushSync();
		expect(places[1]!.getAttribute('aria-selected')).toBe('true');
		expect(
			target.querySelector('[data-pptx-cloud-place="onedrive"] pptx-ui-cloud-files'),
		).not.toBeNull();
		expect(target.querySelector('[data-pptx-search-surface]')).toBeNull();
		const file = { fileName: 'Deck.pptx' } as CloudOpenedFile;
		target
			.querySelector('pptx-ui-cloud-files')!
			.dispatchEvent(
				new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
			);
		expect(open).toHaveBeenCalledWith(file);
		expect(onclose).toHaveBeenCalled();
	});

	it('drops the cloud document before a new presentation is created', () => {
		const clear = vi.fn();
		const oncreatepresentation = vi.fn(() => expect(clear).toHaveBeenCalledOnce());
		const target = render(
			{ providers: [provider], open: vi.fn(), clear },
			{ oncreatepresentation },
		);
		target.querySelector<HTMLButtonElement>('.templates button')!.click();
		expect(oncreatepresentation).toHaveBeenCalledOnce();
	});
});
