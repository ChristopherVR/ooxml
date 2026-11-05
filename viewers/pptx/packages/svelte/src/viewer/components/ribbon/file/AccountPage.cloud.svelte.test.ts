/**
 * File > Account with `fileProviders`: the shared connected services list gets the providers and
 * translated labels, and nothing extra renders without providers.
 */
import type { CloudFileProvider } from 'ooxml-ui/pptx';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { ViewerCloudFiles } from '../../../state/viewer-cloud-files-context';
import AccountPageCloudHarness from './AccountPageCloudHarness.svelte';

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	getLocalStorageUsageSummary: async () => ({ presentationCount: 0, totalBytes: 0 }),
}));

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

function render(cloud: ViewerCloudFiles | undefined): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(AccountPageCloudHarness, { target, props: { cloud } });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

interface ServicesElement extends HTMLElement {
	providers: unknown[];
	labels: Record<string, string>;
}

describe('AccountPage connected services', () => {
	it('renders no connected services without providers', () => {
		const target = render(undefined);
		expect(target.querySelector('pptx-ui-connected-services')).toBeNull();
	});

	it('hands the fileProviders and translated labels to the shared list', () => {
		const target = render({ providers: [provider], open: vi.fn(), clear: vi.fn() });
		const list = target.querySelector<ServicesElement>('pptx-ui-connected-services')!;
		expect(list.providers).toEqual([provider]);
		expect(list.labels['heading']).toBe('Connected services');
		expect(list.labels['connect']).toBe('Connect');
	});
});
