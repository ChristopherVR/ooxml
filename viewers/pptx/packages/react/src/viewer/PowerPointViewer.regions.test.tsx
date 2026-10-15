// @vitest-environment happy-dom
/**
 * Host regions, capabilities and the extra File tab pages, end to end through
 * the real `PowerPointViewer`: a replaced region keeps its slot and hands it to
 * the host, `capabilities.edit` makes the deck read-only live, and the panels
 * and dialogs the host hid never render.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => {
	// One stable translator and i18n object: the File tab refetches its recent
	// files whenever `t` changes identity.
	const t = (key: string, options?: { defaultValue?: string }): string =>
		translationsEn[key] ?? options?.defaultValue ?? key;
	const i18n = {
		language: 'en',
		languages: ['en'],
		options: { resources: { en: {} } },
		changeLanguage: () => Promise.resolve(),
	};
	const value = { t, i18n };
	return { useTranslation: () => value };
});

const { PptxHandler } = await import('pptx-viewer-core');
const { PowerPointViewer } = await import('../index');
type ViewerHandle = import('../index').PowerPointViewerHandle;
type ViewerCustomization = import('../index').ViewerCustomization;
type RegionContext = import('ooxml-ui/pptx').ViewerRegionContext;

async function sampleDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	window.happyDOM?.setViewport({ width: 1600, height: 950 });
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

const tick = (ms = 20) =>
	act(async () => {
		await new Promise((resolve) => {
			setTimeout(resolve, ms);
		});
	});

function ribbonTabs(): HTMLElement[] {
	return Array.from(container.querySelectorAll<HTMLElement>('[role="tab"]'));
}

async function mount(
	customization?: ViewerCustomization,
): Promise<{ current: ViewerHandle | null }> {
	const bytes = await sampleDeck();
	const ref = createRef<ViewerHandle>();
	await act(async () => {
		root.render(
			<PowerPointViewer ref={ref} content={bytes} canEdit customization={customization} />,
		);
	});
	for (let attempt = 0; attempt < 80; attempt += 1) {
		if ((ref.current?.getSlideCount() ?? 0) > 0 && ribbonTabs().length > 0) {
			return ref;
		}
		await tick();
	}
	throw new Error('the viewer never finished loading its deck');
}

/** Open File > `label` (a nav entry of the File tab). */
async function openBackstage(label: string): Promise<void> {
	const fileTab = ribbonTabs().find((tab) => tab.textContent?.trim() === 'File');
	await act(async () => {
		fileTab?.click();
	});
	await tick();
	const entry = Array.from(
		document.body.querySelectorAll<HTMLElement>('[data-pptx-backstage-nav-item]'),
	).find((item) => item.textContent?.trim() === label);
	if (!entry) {
		throw new Error(`no File tab entry named ${label}`);
	}
	await act(async () => {
		entry.click();
	});
	await tick();
}

const SHAPE = { type: 'shape', id: 'probe', x: 0, y: 0, width: 50, height: 50 } as never;

describe('powerPointViewer regions', () => {
	it('mounts a mount function in the inspector slot and updates it on selection', async () => {
		const contexts: RegionContext[] = [];
		const destroy = vi.fn();
		const mountFn = vi.fn((el: HTMLElement, context: RegionContext) => {
			contexts.push(context);
			el.textContent = 'host inspector';
			return { update: (next: RegionContext) => contexts.push(next), destroy };
		});
		const ref = await mount({ regions: { inspector: mountFn } });
		const slot = document.body.querySelector(
			'[data-pptx-inspector] [data-pptx-region="inspector"]',
		);
		expect(slot?.textContent).toBe('host inspector');
		expect(contexts[0]?.region).toBe('inspector');
		expect(contexts[0]?.selectedElementIds).toStrictEqual([]);
		expect(contexts[0]?.canEdit).toBe(true);
		expect(contexts[0]?.slideCount).toBe(1);

		let id: string | undefined;
		await act(async () => {
			id = ref.current!.addElement(SHAPE);
		});
		await tick();
		expect(id).toBeDefined();
		const last = contexts.at(-1)!;
		expect(last.selectedElementIds).toStrictEqual([id]);
		expect(last.viewer.getSlideCount()).toBe(1);

		act(() => {
			root.unmount();
		});
		expect(destroy).toHaveBeenCalledOnce();
		root = createRoot(container);
	});

	it('renders a custom element for dialog.options with a context that can set options', async () => {
		class HostOptions extends HTMLElement {
			context: unknown;
		}
		customElements.define('host-options-view', HostOptions);
		await mount({ regions: { 'dialog.options': 'host-options-view' } });
		await openBackstage('Options');
		const element = document.body.querySelector<HostOptions>('host-options-view');
		expect(element).not.toBeNull();
		const context = element?.context as
			| (RegionContext & { setOption?: unknown; resetOptions?: unknown })
			| undefined;
		expect(context?.region).toBe('dialog.options');
		expect(typeof context?.setOption).toBe('function');
		expect(typeof context?.resetOptions).toBe('function');
		const title = translationsEn['pptx.options.title'];
		// The frame stays; the built-in rail and panes are gone.
		expect(document.body.querySelector(`[role="dialog"][aria-label="${title}"]`)).not.toBeNull();
		expect(document.body.querySelector(`nav[aria-label="${title}"]`)).toBeNull();
	});

	it('still renders a replaced dialog.options when every Options page is hidden', async () => {
		const { VIEWER_OPTIONS_TABS } = await import('ooxml-ui/pptx');
		const mountFn = vi.fn((el: HTMLElement, context: RegionContext) => {
			el.textContent = 'host options';
			el.addEventListener('click', () => context.close());
		});
		await mount({
			options: { hiddenPages: VIEWER_OPTIONS_TABS.map((tab) => tab.id) },
			regions: { 'dialog.options': mountFn },
		});
		await openBackstage('Options');
		const slot = document.body.querySelector<HTMLElement>('[data-pptx-region="dialog.options"]');
		expect(slot?.textContent).toBe('host options');
		// close() behaves like the title-bar X: the dialog goes away.
		await act(async () => {
			slot?.firstElementChild?.dispatchEvent(new Event('click'));
			slot?.dispatchEvent(new Event('click'));
		});
		await tick();
		expect(document.body.querySelector('[data-pptx-region="dialog.options"]')).toBeNull();
	});

	it('lists a host File tab page and mounts its renderer when selected', async () => {
		const mountFn = vi.fn((el: HTMLElement, context: RegionContext) => {
			el.textContent = `page ${context.region}`;
		});
		await mount({
			backstage: { extraPages: [{ id: 'reports', label: 'Reports', render: mountFn }] },
		});
		await openBackstage('Reports');
		expect(mountFn).toHaveBeenCalledOnce();
		const slot = document.body.querySelector('[data-pptx-region="backstage.host.reports"]');
		expect(slot?.textContent).toBe('page backstage.host.reports');
	});
});

describe('powerPointViewer capabilities', () => {
	it('makes the deck read-only with edit: false and restores it live', async () => {
		const ref = await mount();
		await act(async () => {
			ref.current!.setCapability('edit', false);
		});
		await tick();
		expect(ref.current!.addElement(SHAPE)).toBeUndefined();
		await act(async () => {
			ref.current!.setCapability('edit', undefined);
		});
		await tick();
		await act(async () => {
			expect(ref.current!.addElement(SHAPE)).toBeDefined();
		});
	});
});

describe('powerPointViewer hidden panels and dialogs', () => {
	it('never renders the selection pane or the hyperlink dialog the host hid', async () => {
		const ref = await mount({
			hiddenPanels: ['selectionPane'],
			hiddenDialogs: ['hyperlink'],
		});
		expect(ref.current!.getCustomization().hiddenPanels).toStrictEqual(['selectionPane']);
		expect(document.body.querySelector('[data-pptx-selection-pane]')).toBeNull();
		expect(
			document.body.querySelector(`[aria-label="${translationsEn['pptx.hyperlink.title']}"]`),
		).toBeNull();
	});
});
