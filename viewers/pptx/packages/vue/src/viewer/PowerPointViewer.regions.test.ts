import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import type {
	ViewerCustomization,
	ViewerRegionContext,
	ViewerSettingsRegionContext,
} from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose } from './types';

/**
 * Host regions, capabilities and the new hide ids, through the whole viewer
 * (a template that never consults the resolved customisation is what these
 * guard against).
 */
async function mountViewer(
	customization?: ViewerCustomization,
	extra: Record<string, unknown> = {},
): Promise<VueWrapper> {
	const wrapper = mount(PowerPointViewer, {
		props: { content: null, canEdit: true, customization, ...extra },
		attachTo: document.body,
	});
	await flushPromises();
	return wrapper;
}

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
	document.body.innerHTML = '';
});

describe('powerPointViewer host regions', () => {
	it('mounts a regions.inspector function in the inspector slot, updates it and destroys it', async () => {
		const update = vi.fn();
		const destroy = vi.fn();
		const contexts: ViewerRegionContext[] = [];
		const mount_ = vi.fn((container: HTMLElement, context: ViewerRegionContext) => {
			contexts.push(context);
			container.textContent = 'host inspector';
			return {
				update: (next: ViewerRegionContext) => (update(next), contexts.push(next)),
				destroy,
			};
		});
		mounted = await mountViewer(
			{ regions: { inspector: mount_ } },
			{ inspectorInitiallyOpen: true },
		);
		const slot = mounted.get('[data-pptx-inspector] [data-pptx-region="inspector"]');
		expect(slot.text()).toBe('host inspector');
		expect(mounted.find('[data-pptx-chrome="inspector-body"]').exists()).toBeFalsy();
		expect(contexts[0]).toMatchObject({
			region: 'inspector',
			canEdit: true,
			selectedElementIds: [],
			slideCount: 0,
		});
		expect(typeof contexts[0]?.viewer.getSlideCount).toBe('function');
		expect(typeof contexts[0]?.close).toBe('function');

		(mounted.vm as unknown as PowerPointViewerExpose).selectElements(['a']);
		await flushPromises();
		expect(update).toHaveBeenCalledOnce();
		expect(update.mock.calls[0]?.[0].selectedElementIds).toStrictEqual(['a']);
		expect(mount_).toHaveBeenCalledOnce();

		mounted.unmount();
		mounted = undefined;
		expect(destroy).toHaveBeenCalledOnce();
	});

	it('creates a regions[dialog.options] custom element with a settings context', async () => {
		mounted = await mountViewer({ regions: { 'dialog.options': 'host-options-view' } });
		await mounted.get('button[aria-label="Settings"]').trigger('click');
		await flushPromises();
		const element = document.body.querySelector<HTMLElement & { context?: unknown }>(
			'host-options-view',
		);
		expect(element).not.toBeNull();
		const context = element?.context as ViewerSettingsRegionContext;
		expect(context.region).toBe('dialog.options');
		expect(typeof context.setOption).toBe('function');
		expect(typeof context.resetOptions).toBe('function');
		expect(context.options.general).toBeDefined();
		// The frame stays; the built-in rail and footer are gone.
		const dialog = document.body.querySelector('pptx-ui-options-dialog');
		expect(dialog?.hasAttribute('custom-body')).toBe(true);
		expect(element?.parentElement?.slot).toBe('body');
		expect(dialog?.shadowRoot?.querySelector('[role="tablist"]')).toBeNull();
		expect(dialog?.shadowRoot?.querySelector('[data-action="ok"]')).toBeNull();

		context.close();
		await flushPromises();
		expect(document.body.querySelector('host-options-view')).toBeNull();
	});

	it('routes setOption through the options store and respects locked settings', async () => {
		mounted = await mountViewer({
			regions: { 'dialog.options': 'host-options-view' },
			options: { locked: { 'general.enableCustomFontUpload': false } },
		});
		await mounted.get('button[aria-label="Settings"]').trigger('click');
		await flushPromises();
		const element = document.body.querySelector<HTMLElement & { context?: unknown }>(
			'host-options-view',
		);
		const read = () => (element?.context as ViewerSettingsRegionContext).options;
		const before = read().general.showStartScreen;
		(element?.context as ViewerSettingsRegionContext).setOption(
			'general.enableCustomFontUpload',
			true,
		);
		await flushPromises();
		expect(read().general.enableCustomFontUpload).toBe(false);
		(element?.context as ViewerSettingsRegionContext).setOption('general.showStartScreen', !before);
		await flushPromises();
		expect(read().general.showStartScreen).toBe(!before);
	});

	it('shows host File tab pages in the nav and mounts their renderer on selection', async () => {
		const renderer = vi.fn((container: HTMLElement, context: ViewerRegionContext) => {
			container.textContent = `host page ${context.region}`;
		});
		mounted = await mountViewer({
			backstage: { extraPages: [{ id: 'reports', label: 'Reports', render: renderer }] },
		});
		const fileTab = mounted.findAll('[role="tab"]').find((tab) => tab.text().trim() === 'File');
		await fileTab!.trigger('click');
		await flushPromises();
		const navItem = mounted
			.findAll('[data-pptx-backstage-nav-item]')
			.find((item) => item.text().trim() === 'Reports');
		expect(navItem).toBeDefined();
		await navItem!.trigger('click');
		await flushPromises();
		expect(renderer).toHaveBeenCalledOnce();
		expect(renderer.mock.calls[0]?.[1].region).toBe('backstage.host.reports');
		expect(mounted.get('[data-pptx-region="backstage.host.reports"]').text()).toBe(
			'host page backstage.host.reports',
		);
	});
});

describe('powerPointViewer capabilities and new hide ids', () => {
	it('is read-only with capabilities.edit false and restores live', async () => {
		mounted = await mountViewer({ capabilities: { edit: false } });
		const viewer = mounted.vm as unknown as PowerPointViewerExpose;
		expect(mounted.find('main.is-editable').exists()).toBeFalsy();
		viewer.setCapability('edit', undefined);
		await flushPromises();
		expect(mounted.find('main.is-editable').exists()).toBeTruthy();
		viewer.setCapability('edit', false);
		await flushPromises();
		expect(mounted.find('main.is-editable').exists()).toBeFalsy();
	});

	it('carries the ribbon control ids the new hide ids rely on', async () => {
		mounted = await mountViewer();
		for (const id of ['home.editing.find', 'home.editing.replace']) {
			expect(mounted.find(`[data-ribbon-control="${id}"]`).exists(), id).toBeTruthy();
		}
	});

	// Hiding an id removes its entry points, but the panel must also refuse to render
	// when opened through any other path (here: the viewer's own open flags).
	it.each([
		['findReplace panel', { hiddenPanels: ['findReplace'] }, '[role="search"]', 'findOpen'],
		[
			'passwordProtection dialog',
			{ hiddenDialogs: ['passwordProtection'] },
			'#pptx-vue-password-error, [data-pptx-password-dialog], input[type="password"]',
			'password',
		],
	] as const)('does not render a hidden %s', async (_name, hidden, selector, flag) => {
		const open = (wrapper: VueWrapper): void => {
			const state = wrapper.vm as unknown as Record<
				string,
				{ showPasswordDialog?: { value: boolean } }
			>;
			if (flag === 'password') {
				state[flag]!.showPasswordDialog!.value = true;
			} else {
				(wrapper.vm as unknown as Record<string, boolean>)[flag] = true;
			}
		};
		mounted = await mountViewer();
		open(mounted);
		await flushPromises();
		expect(document.body.querySelector(selector)).not.toBeNull();
		mounted.unmount();

		mounted = await mountViewer(hidden as ViewerCustomization);
		open(mounted);
		await flushPromises();
		expect(document.body.querySelector(selector)).toBeNull();
	});
});
