/**
 * File > Account with `fileProviders`: the shared connected services list appears with the
 * providers and translated labels, and nothing extra renders without providers.
 */
import { mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import type { CloudFileProvider } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { ViewerCloudFilesKey } from '../../composables/viewer-cloud-files';
import type { ViewerCloudFiles } from '../../composables/viewer-cloud-files';
import AccountPage from './AccountPage.vue';

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	getLocalStorageUsageSummary: async () => ({ presentationCount: 0, totalBytes: 0 }),
}));

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

let wrapper: VueWrapper | undefined;
afterEach(() => {
	wrapper?.unmount();
	wrapper = undefined;
});

describe('AccountPage connected services', () => {
	it('renders no connected services without providers', () => {
		wrapper = mount(AccountPage, { attachTo: document.body });
		expect(wrapper.find('pptx-ui-connected-services').exists()).toBe(false);
	});

	it('hands the fileProviders and translated labels to the shared list', () => {
		const cloud: ViewerCloudFiles = {
			providers: () => [provider],
			open: vi.fn(),
			clear: vi.fn(),
		};
		wrapper = mount(AccountPage, {
			attachTo: document.body,
			global: { provide: { [ViewerCloudFilesKey as symbol]: cloud } },
		});
		const element = wrapper.get('pptx-ui-connected-services').element as unknown as {
			providers: unknown[];
			labels: Record<string, string>;
		};
		expect(element.providers).toStrictEqual([provider]);
		expect(element.labels.heading).toBe('Connected services');
		expect(element.labels.connect).toBe('Connect');
	});
});
