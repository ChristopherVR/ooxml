// @vitest-own-worker: mounts several full viewers; see region-customization.test-support.ts.
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	destroyMounted,
	internals,
	mount,
	region,
	storeOf,
} from './region-customization.test-support';
import type { ViewerRegionContext } from './region-customization.test-support';

afterEach(destroyMounted);

describe('vanilla replaced regions: inspector', () => {
	it('mounts, updates on selection change and destroys with the viewer', () => {
		const update = vi.fn();
		const destroy = vi.fn();
		const mountRegion = vi.fn((container: HTMLElement, _context: ViewerRegionContext) => {
			container.textContent = 'host inspector';
			return { update, destroy };
		});
		const { container, viewer } = mount({
			editable: true,
			customization: { regions: { inspector: mountRegion } },
		});
		const slot = region(container, 'inspector');
		expect(slot?.textContent).toBe('host inspector');
		expect(slot?.closest('[data-pptx-inspector]')).not.toBeNull();
		const first = mountRegion.mock.calls[0]?.[1] as ViewerRegionContext;
		expect(first.region).toBe('inspector');
		// The chrome is built before the editable flag lands; the update delivers it.
		expect((update.mock.calls.at(-1)?.[0] as ViewerRegionContext).canEdit).toBe(true);
		expect(first.selectedElementIds).toStrictEqual([]);
		storeOf(viewer).set({ selectedElementIds: ['a'], selectedElementId: 'a' });
		expect(update).toHaveBeenCalled();
		const latest = update.mock.calls.at(-1)?.[0] as ViewerRegionContext;
		expect(latest.selectedElementIds).toStrictEqual(['a']);
		viewer.destroy();
		expect(destroy).toHaveBeenCalledOnce();
	});

	it('keeps a live host (and its container) across an unrelated chrome rebuild', () => {
		const destroy = vi.fn();
		const mountRegion = vi.fn((container: HTMLElement) => {
			container.textContent = 'host';
			return { destroy };
		});
		const { container, viewer } = mount({
			editable: true,
			customization: { regions: { inspector: mountRegion } },
		});
		const before = region(container, 'inspector');
		viewer.hideRibbonTab('draw');
		const after = region(container, 'inspector');
		expect(mountRegion).toHaveBeenCalledOnce();
		expect(destroy).not.toHaveBeenCalled();
		expect(after).toBe(before);
		expect(after?.isConnected).toBe(true);
		expect(after?.closest('[data-pptx-inspector]')).not.toBeNull();
		// A changed renderer destroys the old host; clearing the region destroys the new one.
		const next = vi.fn(() => ({ destroy }));
		viewer.setRegion('inspector', next);
		expect(destroy).toHaveBeenCalledOnce();
		expect(next).toHaveBeenCalledOnce();
		viewer.setRegion('inspector', undefined);
		expect(destroy).toHaveBeenCalledTimes(2);
	});

	it('restores the built-in pane when the region is cleared', () => {
		const { container, viewer } = mount({
			editable: true,
			customization: {
				regions: { inspector: (el: HTMLElement) => void (el.textContent = 'host') },
			},
		});
		expect(region(container, 'inspector')).not.toBeNull();
		viewer.setRegion('inspector', undefined);
		expect(region(container, 'inspector')).toBeNull();
	});
});

describe('vanilla replaced regions: dialogs', () => {
	it('renders dialog.options as a custom element with a settable context', () => {
		class OptionsView extends HTMLElement {
			context: ViewerRegionContext | undefined;
		}
		customElements.define('test-options-view', OptionsView);
		const { viewer } = mount({
			editable: true,
			customization: { regions: { 'dialog.options': 'test-options-view' } },
		});
		internals(viewer).openSettings();
		const element = document.querySelector<OptionsView>('test-options-view');
		expect(element).not.toBeNull();
		expect(element?.context?.region).toBe('dialog.options');
		expect(typeof (element?.context as { setOption?: unknown }).setOption).toBe('function');
		// The built-in rail and panes are not built; the frame is.
		expect(document.querySelector('.pptxv-options-nav')).toBeNull();
		expect(document.querySelector('.pptxv-parity-dialog')).not.toBeNull();
		// Escape dismisses the frame and releases the renderer.
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(document.querySelector('test-options-view')).toBeNull();
	});

	it('routes setOption through the options store and respects locks', () => {
		let context: ViewerRegionContext | undefined;
		const { viewer } = mount({
			editable: true,
			customization: {
				regions: {
					'dialog.options': (_el: HTMLElement, ctx: ViewerRegionContext) => {
						context = ctx;
					},
				},
			},
		});
		viewer.lockSetting('general.userInitials', 'ZZ');
		internals(viewer).openSettings();
		expect(context).toBeDefined();
		const settings = context as unknown as {
			setOption(id: string, value: unknown): void;
			options: { general: { userInitials: string } };
		};
		settings.setOption('general.userInitials', 'AB');
		expect(settings.options.general.userInitials).toBe('ZZ');
	});

	it('destroys an open dialog renderer on viewer destroy', () => {
		const destroy = vi.fn();
		const { viewer } = mount({
			editable: true,
			customization: { regions: { 'dialog.print': () => destroy } },
		});
		internals(viewer).openPrintDialog();
		expect(document.querySelector('[data-pptx-region="dialog.print"]')).not.toBeNull();
		viewer.destroy();
		expect(destroy).toHaveBeenCalledOnce();
	});
});
