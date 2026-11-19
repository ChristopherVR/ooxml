import { mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { resolveCustomization } from 'ooxml-ui/pptx';
import type { ToolbarActionId, ViewerCustomization } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';
import { shallowRef } from 'vue';

import { ViewerCustomizationKey } from '../../composables/useViewerCustomization';
import FileSection from './FileSection.vue';

/** Click the "Export" backstage-nav button (it has no aria-label, only its text). */
async function openExportPage(wrapper: VueWrapper): Promise<void> {
	const button = wrapper.findAll('nav button').find((btn) => btn.text().includes('Export'));
	if (!button) {
		throw new Error('Export nav button not found');
	}
	await button.trigger('click');
}

const BASE_PROPS = {
	onClose: () => {},
	onCreatePresentation: () => {},
	onExportPng: () => {},
	onExportPdf: () => {},
	onExportJson: () => {},
	onExportVideo: () => {},
	onExportGif: () => {},
	onSaveAsPptx: () => {},
	onSaveAsPpsx: () => {},
	onSaveAsPptm: () => {},
	onSaveAsPpt: () => {},
	hasMacros: false,
	onCopySlideAsImage: () => {},
	onPrint: () => {},
};

function mountFileSection(hiddenActions?: ToolbarActionId[]) {
	return mount(FileSection, { props: { ...BASE_PROPS, hiddenActions } });
}

/**
 * FileSection: the File-tab backstage. Covers the `hiddenActions` gating
 * added for issue #64: navigating to the Export page shows its action cards
 * by default, and shows none when the host hides the shared 'export' id.
 */
describe('fileSection', () => {
	it('shows the Export page action cards by default', async () => {
		const wrapper = mountFileSection(undefined);
		await openExportPage(wrapper);
		expect(wrapper.text()).toContain('Create PDF');
	});

	// React parity: hiding the 'export' action drops the Export nav entry
	// itself (it used to stay, leading to an empty page), not just its cards.
	it('drops the Export page from the nav when "export" is hidden', () => {
		const wrapper = mountFileSection(['export']);
		expect(navLabels(wrapper)).not.toContain('Export');
		expect(wrapper.text()).not.toContain('Create PDF');
	});
});

describe('fileSection host containment', () => {
	it('sits on the viewer host box and closes on Escape pressed inside it', () => {
		const outside = document.createElement('button');
		const host = document.createElement('div');
		host.setAttribute('data-pptx-theme-root', '');
		document.body.append(outside, host);
		vi.spyOn(host, 'getBoundingClientRect').mockReturnValue({
			left: 160,
			top: 40,
			width: 1000,
			height: 700,
		} as DOMRect);
		const onClose = vi.fn();
		const wrapper = mount(FileSection, { props: { ...BASE_PROPS, onClose }, attachTo: host });
		const panel = wrapper.get('[role="dialog"]').element as HTMLElement;
		expect([
			panel.style.left,
			panel.style.top,
			panel.style.width,
			panel.style.height,
		]).toStrictEqual(['160px', '40px', '1000px', '700px']);
		const escape = (target: Element) =>
			target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		escape(outside);
		expect(onClose).not.toHaveBeenCalled();
		escape(panel.querySelector('button')!);
		expect(onClose).toHaveBeenCalledOnce();
		wrapper.unmount();
		host.remove();
		outside.remove();
		vi.restoreAllMocks();
	});
});

function navLabels(wrapper: VueWrapper): string[] {
	return wrapper.findAll('nav button').map((btn) => btn.text().trim());
}

function mountCustomized(customization: ViewerCustomization) {
	return mount(FileSection, {
		props: BASE_PROPS,
		global: {
			provide: {
				[ViewerCustomizationKey as symbol]: shallowRef(resolveCustomization(customization)),
			},
		},
	});
}

describe('fileSection under UI customization', () => {
	it('drops hidden backstage pages from the nav', () => {
		const wrapper = mountCustomized({ backstage: { hiddenPages: ['open', 'account'] } });
		const labels = navLabels(wrapper);
		expect(labels).toContain('New');
		expect(labels).not.toContain('Open');
		expect(labels).not.toContain('Account');
	});

	it('drops hidden cards from a page', async () => {
		const wrapper = mountCustomized({ backstage: { hiddenCards: ['pdf'] } });
		await openExportPage(wrapper);
		expect(wrapper.text()).not.toContain('Create PDF');
		expect(wrapper.text()).toContain('Create an Animated GIF');
	});

	it('drops the Options page when the host removes the Options dialog', () => {
		const wrapper = mountCustomized({ hiddenDialogs: ['options'] });
		expect(navLabels(wrapper)).not.toContain('Options');
	});
});

describe('fileSection with host-owned storage', () => {
	it('runs onSave from the Save nav item, not the Save As .pptx handler', async () => {
		const calls: string[] = [];
		const wrapper = mount(FileSection, {
			props: {
				...BASE_PROPS,
				onSave: () => calls.push('save'),
				onSaveAsPptx: () => calls.push('saveAs'),
			},
		});
		const save = wrapper.findAll('nav button').find((btn) => btn.text().trim() === 'Save');
		await save!.trigger('click');
		expect(calls).toStrictEqual(['save']);
	});
});
