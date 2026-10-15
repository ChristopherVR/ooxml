// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { registerPptxWebControls } from '../../web-components';
import type {
	ViewerRegionContext,
	ViewerSettingsRegionContext,
} from '../customization/customization-regions';
import { resolveCustomization } from '../customization/customization-resolve';
import type { ViewerCustomization } from '../customization/customization-types';
import {
	bindPptxOptionsDialog,
	type PptxOptionsDialogElement,
	type PptxOptionsDialogHost,
} from './options-dialog-binding';
import { DEFAULT_VIEWER_OPTIONS } from './viewer-options';

const t = (key: string) => key;

const BASE: ViewerRegionContext = {
	region: 'inspector',
	viewer: {} as ViewerRegionContext['viewer'],
	mode: 'edit',
	canEdit: true,
	activeSlideIndex: 0,
	slideCount: 3,
	selectedElementIds: [],
	locale: 'en',
	close: () => {},
};

describe('bindPptxOptionsDialog with a dialog.options region', () => {
	beforeAll(() => registerPptxWebControls());
	afterEach(() => document.body.replaceChildren());

	function mount(customization: ViewerCustomization) {
		const dialog = document.createElement('pptx-ui-options-dialog') as PptxOptionsDialogElement;
		document.body.append(dialog);
		const host = {
			setOption: vi.fn(),
			restoreOptions: vi.fn(),
			resetOptions: vi.fn(),
			close: vi.fn(),
		} satisfies PptxOptionsDialogHost;
		const binding = bindPptxOptionsDialog(dialog, host);
		const input = {
			options: DEFAULT_VIEWER_OPTIONS,
			t,
			customization: resolveCustomization(customization),
			aiEnabled: false,
			regionContext: BASE,
		};
		binding.sync(input);
		return { dialog, host, binding, input };
	}

	it('mounts the host renderer in the body with the options context', () => {
		let context: ViewerSettingsRegionContext | undefined;
		const { dialog } = mount({
			regions: {
				'dialog.options': (container, ctx) => {
					context = ctx as ViewerSettingsRegionContext;
					container.textContent = 'Host settings';
				},
			},
		});
		expect(dialog.customBody).toBe(true);
		const container = dialog.querySelector<HTMLElement>('[data-pptx-region="dialog.options"]');
		expect(container?.slot).toBe('body');
		expect(container?.textContent).toBe('Host settings');
		expect(context?.region).toBe('dialog.options');
		expect(context?.options).toBe(DEFAULT_VIEWER_OPTIONS);
		expect(context?.slideCount).toBe(3);
	});

	it('routes setOption, resetOptions and close through the host, honouring locks', () => {
		let context: ViewerSettingsRegionContext | undefined;
		const { host } = mount({
			options: { locked: { 'general.userName': 'Kiosk' } },
			regions: {
				'dialog.options': (_container, ctx) => {
					context = ctx as ViewerSettingsRegionContext;
				},
			},
		});
		context!.setOption('general.showMiniToolbar', false);
		context!.setOption('general.userName', 'Someone else');
		context!.resetOptions();
		context!.close();
		expect(host.setOption).toHaveBeenCalledTimes(1);
		expect(host.setOption).toHaveBeenCalledWith('general', 'showMiniToolbar', false);
		expect(host.resetOptions).toHaveBeenCalledTimes(1);
		expect(host.close).toHaveBeenCalledTimes(1);
		expect(host.restoreOptions).not.toHaveBeenCalled();
	});

	it('updates without remounting and falls back to the built-in body when removed', () => {
		const destroy = vi.fn();
		const update = vi.fn();
		const { dialog, binding, input } = mount({
			regions: { 'dialog.options': () => ({ update, destroy }) },
		});
		const options = {
			...DEFAULT_VIEWER_OPTIONS,
			general: { ...DEFAULT_VIEWER_OPTIONS.general, userName: 'Ada' },
		};
		binding.sync({ ...input, options });
		expect(update).toHaveBeenCalledTimes(1);
		expect(destroy).not.toHaveBeenCalled();
		binding.sync({ ...input, customization: resolveCustomization(undefined) });
		expect(destroy).toHaveBeenCalledTimes(1);
		expect(dialog.customBody).toBe(false);
		expect(dialog.querySelector('[data-pptx-region]')).toBeNull();
	});

	it('keeps the built-in body without a region context and cleans up on dispose', () => {
		const mountFn = vi.fn();
		const { dialog, binding, input } = mount({ regions: { 'dialog.options': mountFn } });
		binding.sync({ ...input, regionContext: null });
		expect(dialog.customBody).toBe(false);
		expect(dialog.querySelector('[data-pptx-region]')).toBeNull();
		binding.sync(input);
		expect(mountFn).toHaveBeenCalledTimes(2);
		binding.dispose();
		expect(dialog.querySelector('[data-pptx-region]')).toBeNull();
	});
});
