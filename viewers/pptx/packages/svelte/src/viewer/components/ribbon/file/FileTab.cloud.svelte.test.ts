/**
 * File > Open with cloud places: the Open page renders the shared places element with the
 * providers and the translated labels, and hands a file opened there to the viewer.
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

interface PlacesElement extends HTMLElement {
	providers: unknown[];
	accept: string[];
	labels: Record<string, string>;
}

describe('FileTab open places', () => {
	it('keeps the Open page as it was without providers', () => {
		const target = render(undefined);
		openPage(target);
		const element = target.querySelector<PlacesElement>('pptx-ui-open-places')!;
		expect(element.providers).toEqual([]);
		expect(target.querySelector('[data-pptx-search-input]')).not.toBeNull();
		expect(target.textContent).toContain('Browse');
	});

	it('hands the providers to the places element and opens a provider file', () => {
		const open = vi.fn();
		const onclose = vi.fn();
		const target = render({ providers: [provider], open, clear: vi.fn() }, { onclose });
		openPage(target);
		const element = target.querySelector<PlacesElement>('pptx-ui-open-places')!;
		expect(element.providers).toEqual([provider]);
		expect(element.accept).toEqual(['pptx']);
		expect(element.labels.thisDevice).toBe('This device');
		expect(element.labels.places).toBe('Places');
		const file = { fileName: 'Deck.pptx' } as CloudOpenedFile;
		element.dispatchEvent(
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
