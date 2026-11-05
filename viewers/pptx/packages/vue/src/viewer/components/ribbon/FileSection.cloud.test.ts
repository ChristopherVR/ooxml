/**
 * File > Open with cloud places: the Open page lists "This device" and one place per provider,
 * shows the provider's file browser, and hands a file opened there to the viewer.
 */
import { mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { ViewerCloudFilesKey } from '../../composables/viewer-cloud-files';
import type { ViewerCloudFiles } from '../../composables/viewer-cloud-files';
import FileSection from './FileSection.vue';

const noop = () => {};
const BASE_PROPS = {
	onClose: noop,
	onCreatePresentation: noop,
	onExportPng: noop,
	onExportPdf: noop,
	onExportJson: noop,
	onExportVideo: noop,
	onExportGif: noop,
	onSaveAsPptx: noop,
	onSaveAsPpsx: noop,
	onSaveAsPptm: noop,
	onSaveAsPpt: noop,
	hasMacros: false,
	onCopySlideAsImage: noop,
	onPrint: noop,
};

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

async function openPage(view: VueWrapper): Promise<void> {
	const open = view
		.findAll('[data-pptx-backstage-nav-item]')
		.find((button) => button.text().trim() === 'Open');
	await open!.trigger('click');
}

describe('fileSection open places', () => {
	it('keeps the Open page as it was without providers', async () => {
		wrapper = mount(FileSection, { props: BASE_PROPS, attachTo: document.body });
		await openPage(wrapper);
		expect(wrapper.find('[data-pptx-open-place]').exists()).toBe(false);
		expect(wrapper.text()).toContain('Browse');
	});

	it('lists this device and the provider, and opens a provider file', async () => {
		const cloud: ViewerCloudFiles = {
			providers: () => [provider],
			open: vi.fn(),
			clear: vi.fn(),
		};
		const onClose = vi.fn();
		wrapper = mount(FileSection, {
			props: { ...BASE_PROPS, onClose },
			attachTo: document.body,
			global: { provide: { [ViewerCloudFilesKey as symbol]: cloud } },
		});
		await openPage(wrapper);
		const tablist = wrapper.get('[role="tablist"]');
		expect(tablist.attributes('aria-label')).toBe('Places');
		const places = wrapper.findAll('[data-pptx-open-place]');
		expect(places.map((place) => place.text().trim())).toStrictEqual(['This device', 'OneDrive']);
		expect(places[0]!.attributes('aria-selected')).toBe('true');
		expect(wrapper.find('pptx-ui-cloud-files').exists()).toBe(false);
		await places[1]!.trigger('click');
		expect(wrapper.find('[data-pptx-cloud-place="onedrive"]').exists()).toBe(true);
		expect(wrapper.find('[data-pptx-search-input]').exists()).toBe(false);
		const file = { fileName: 'Deck.pptx' } as CloudOpenedFile;
		wrapper
			.get('pptx-ui-cloud-files')
			.element.dispatchEvent(
				new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
			);
		expect(cloud.open).toHaveBeenCalledWith(file);
		expect(onClose).toHaveBeenCalled();
	});

	it('clears the cloud file before creating a new presentation', async () => {
		const calls: string[] = [];
		const cloud: ViewerCloudFiles = {
			providers: () => [provider],
			open: vi.fn(),
			clear: () => calls.push('clear'),
		};
		wrapper = mount(FileSection, {
			props: { ...BASE_PROPS, onCreatePresentation: () => calls.push('create') },
			global: { provide: { [ViewerCloudFilesKey as symbol]: cloud } },
		});
		await wrapper.find('.grid button').trigger('click');
		expect(calls).toStrictEqual(['clear', 'create']);
	});
});
