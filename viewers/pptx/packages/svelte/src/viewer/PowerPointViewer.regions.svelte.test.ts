import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type {
	PowerPointViewerAPI,
	ViewerCustomization,
	ViewerCustomizationApi,
	ViewerRegionContext,
	ViewerSettingsRegionContext,
} from 'ooxml-ui/pptx';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import SlideTemplatesLauncher from './components/ribbon/home/SlideTemplatesLauncher.svelte';
import ReviewTab from './components/ribbon/review/ReviewTab.svelte';
import type { EditorState } from './editor/editor-state.svelte';
import PowerPointViewer from './PowerPointViewer.svelte';
import {
	ViewerCustomizationState,
	customizationContext,
} from './state/viewer-customization.svelte';
import type { PowerPointViewerProps } from './types';

/**
 * Host-replaced regions, host File tab pages and the `edit` capability, through
 * the public component and its `customization` input.
 */
registerPptxWebControls();
const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

type ViewerInstance = ViewerCustomizationApi &
	Pick<PowerPointViewerAPI, 'getElements' | 'selectElements' | 'clearSelection'>;

async function mountViewer(
	customization: ViewerCustomization,
	props: Partial<PowerPointViewerProps> = {},
): Promise<{ target: HTMLElement; instance: ViewerInstance }> {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const onload = vi.fn();
	const merged = Object.assign(props, {
		source: new Uint8Array(readFileSync(FIXTURE)),
		editable: props.editable ?? true,
		inspectorInitiallyOpen: true,
		customization,
		onload,
	});
	const instance = mount(PowerPointViewer, { target, props: merged }) as unknown as ViewerInstance;
	flushSync();
	cleanup = () => {
		unmount(instance as unknown as Record<string, unknown>);
		target.remove();
	};
	await vi.waitFor(() => expect(onload).toHaveBeenCalledOnce(), { timeout: 15000 });
	flushSync();
	return { target, instance };
}

function openFileTab(target: HTMLElement): void {
	const file = [
		...target.querySelectorAll<HTMLElement>('.pptx-svelte-ribbon-tabs [role="tab"]'),
	].find((tab) => tab.textContent?.trim() === 'File');
	file!.click();
	flushSync();
}

function navItem(target: HTMLElement, label: string): HTMLElement | undefined {
	return [...target.querySelectorAll<HTMLElement>('[data-pptx-backstage-nav-item]')].find(
		(item) => item.textContent?.trim() === label,
	);
}

describe('host regions in the svelte viewer', () => {
	it('mounts a function renderer in the inspector slot, updates it on selection and destroys it', async () => {
		const update = vi.fn();
		const destroy = vi.fn();
		const mountFn = vi.fn((container: HTMLElement, context: ViewerRegionContext) => {
			container.textContent = `host inspector ${context.selectedElementIds.length}`;
			return { update, destroy };
		});
		const { target, instance } = await mountViewer({ regions: { inspector: mountFn } });
		const host = target.querySelector('[data-pptx-inspector] [data-pptx-region="inspector"]');
		expect(host?.textContent).toBe('host inspector 0');
		// The built-in pane content is gone, the column stays.
		expect(target.querySelector('.pptx-svelte-inspector-tabs')).toBeNull();
		const first = mountFn.mock.calls[0]![1];
		expect(first.region).toBe('inspector');
		expect(first.canEdit).toBe(true);
		expect(typeof first.viewer.getElements).toBe('function');

		const id = instance.getElements()[0]!.id;
		instance.selectElements([id]);
		flushSync();
		const latest = update.mock.calls.at(-1)![0] as ViewerRegionContext;
		expect(latest.selectedElementIds).toStrictEqual([id]);
		expect(mountFn).toHaveBeenCalledOnce();

		cleanup?.();
		cleanup = undefined;
		expect(destroy).toHaveBeenCalledOnce();
	});

	it('gives regions the full public handle as context.viewer', async () => {
		const mountFn = vi.fn((container: HTMLElement, _context: ViewerRegionContext) => {
			container.textContent = 'host';
		});
		await mountViewer({ regions: { inspector: mountFn } });
		const context = mountFn.mock.calls[0]![1];
		const viewer = context.viewer as unknown as Record<string, unknown>;
		for (const name of [
			'getSlideCount',
			'getContent',
			'undo',
			'canUndo',
			'setRegion',
			'setCapability',
			'hideRibbonTab',
			'getCustomization',
		]) {
			expect(typeof viewer[name], name).toBe('function');
		}
		expect((context.viewer as PowerPointViewerAPI).getSlideCount()).toBe(context.slideCount);
	});

	it('replaces the whole notes panel inside a content-sized slot', async () => {
		const update = vi.fn();
		const notes = vi.fn((container: HTMLElement, _context: ViewerRegionContext) => {
			container.textContent = 'host notes';
			return { update, destroy: vi.fn() };
		});
		const { target } = await mountViewer({ regions: { notes } });
		const slot = target.querySelector<HTMLElement>('.pptx-svelte-notes-region');
		expect(slot?.textContent).toBe('host notes');
		expect(getComputedStyle(slot!).height).not.toBe('160px');
		expect(slot!.querySelector('.pptx-svelte-notes-header')).toBeNull();
		expect(notes.mock.calls[0]![1].canEdit).toBe(true);
	});

	it('hosts a custom element for dialog.options with a settings context and no built-in panes', async () => {
		const tag = 'pptx-svelte-test-options-host';
		const contexts: ViewerSettingsRegionContext[] = [];
		if (!customElements.get(tag)) {
			customElements.define(
				tag,
				class extends HTMLElement {
					set context(value: ViewerSettingsRegionContext) {
						contexts.push(value);
					}
				},
			);
		}
		const { target } = await mountViewer({ regions: { 'dialog.options': tag } });
		openFileTab(target);
		navItem(target, 'Options')!.click();
		flushSync();
		const element = target.querySelector(`[data-pptx-region="dialog.options"] ${tag}`);
		expect(element).not.toBeNull();
		const context = contexts.at(-1)!;
		expect(context.region).toBe('dialog.options');
		expect(typeof context.setOption).toBe('function');
		expect(typeof context.resetOptions).toBe('function');
		expect(context.options).toBeDefined();
		const dialog = target.querySelector('pptx-ui-options-dialog');
		expect(dialog?.hasAttribute('custom-body')).toBe(true);
		expect(dialog?.shadowRoot?.querySelector('[role="tablist"]')).toBeNull();
		// The frame stays: close dismisses the dialog.
		context.close();
		flushSync();
		expect(target.querySelector('[data-pptx-region="dialog.options"]')).toBeNull();
	});

	it('shows a host File tab page in the nav and mounts its renderer', async () => {
		const mountFn = vi.fn((container: HTMLElement) => {
			container.textContent = 'host page body';
		});
		const { target } = await mountViewer({
			backstage: { extraPages: [{ id: 'reports', label: 'Reports', render: mountFn }] },
		});
		openFileTab(target);
		const item = navItem(target, 'Reports');
		expect(item).toBeDefined();
		item!.click();
		flushSync();
		const region = target.querySelector('[data-pptx-region="backstage.host.reports"]');
		expect(region?.textContent).toBe('host page body');
		expect(target.querySelector('.bs h1')?.textContent).toBe('Reports');
	});

	it('makes the viewer read-only with capabilities.edit false and restores it live', async () => {
		const { target, instance } = await mountViewer({ capabilities: { edit: false } });
		expect(target.querySelector('[data-pptx-inspector]')).toBeNull();
		instance.setCapability('edit', undefined);
		flushSync();
		expect(target.querySelector('[data-pptx-inspector]')).not.toBeNull();
		instance.setCapability('edit', false);
		flushSync();
		expect(target.querySelector('[data-pptx-inspector]')).toBeNull();
	});
});

describe('hidden panels and dialogs gate their render', () => {
	function reviewTab(hidden: boolean): HTMLElement {
		const target = document.createElement('div');
		document.body.appendChild(target);
		const state = new ViewerCustomizationState(
			hidden ? { hiddenPanels: ['accessibilityPane'] } : {},
		);
		const instance = mount(ReviewTab, {
			target,
			props: { slides: [], onnavigate: vi.fn() },
			context: customizationContext(state),
		});
		flushSync();
		cleanup = () => {
			unmount(instance);
			target.remove();
			state.destroy();
		};
		const button = [...target.querySelectorAll('pptx-ui-ribbon-command')]
			.map((host) => host.shadowRoot!.querySelector<HTMLButtonElement>('button')!)
			.find((candidate) => candidate.textContent?.includes('Accessibility'));
		button!.click();
		flushSync();
		return target;
	}

	it('renders the accessibility pane unless it is hidden', () => {
		expect(reviewTab(false).querySelector('.pptx-svelte-review-panel')).not.toBeNull();
		cleanup?.();
		expect(reviewTab(true).querySelector('.pptx-svelte-review-panel')).toBeNull();
	});

	it('does not render the slide templates dialog when it is hidden', () => {
		for (const [hidden, expected] of [
			[false, true],
			[true, false],
		] as const) {
			const target = document.createElement('div');
			document.body.appendChild(target);
			const state = new ViewerCustomizationState(
				hidden ? { hiddenDialogs: ['slideTemplates'] } : {},
			);
			const instance = mount(SlideTemplatesLauncher, {
				target,
				props: {
					editor: { editable: true } as EditorState,
					onnavigate: () => undefined,
					open: true,
				},
				context: customizationContext(state),
			});
			flushSync();
			expect(target.querySelector('[role="dialog"], dialog') !== null).toBe(expected);
			unmount(instance);
			target.remove();
			state.destroy();
		}
	});
});
