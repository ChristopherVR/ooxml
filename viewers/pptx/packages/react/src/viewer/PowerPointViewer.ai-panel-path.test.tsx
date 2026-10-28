// @vitest-environment happy-dom
/**
 * Which config the real `PowerPointViewer` gives its chat panel for each kind of `ai` value.
 * The lazy chat panel is replaced by a recorder, so nothing loads the real panel chunk:
 * a `createAiAssistant` handle passes its own config, a plain config passes itself, and a
 * hand-built object is a plain config, never a handle. Handles come only from the factory.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createAiAssistant } from '../ai';
import type { AiChatPanelLazyProps } from './components/ai/AiChatPanelLazy';

const lazyPanel = vi.hoisted(() => ({
	received: [] as AiChatPanelLazyProps[],
	mounts: 0,
}));

// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock<typeof import('./components/ai/AiChatPanelLazy')>(
	import('./components/ai/AiChatPanelLazy'),
	() => ({
		AiChatPanelLazy(props: AiChatPanelLazyProps) {
			lazyPanel.received.push(props);
			useEffect(() => {
				lazyPanel.mounts += 1;
			}, []);
			return (
				<button type='button' data-testid='lazy-panel' onClick={props.onClose}>
					close
				</button>
			);
		},
	}),
);

// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
		i18n: {
			language: 'en',
			languages: ['en'],
			options: { resources: { en: {} } },
			changeLanguage: () => Promise.resolve(),
		},
	}),
}));

const { PptxHandler } = await import('pptx-viewer-core');
const { PowerPointViewer } = await import('../index');
type ViewerHandle = import('../index').PowerPointViewerHandle;
type ViewerProps = import('../index').PowerPointViewerProps;

const config: PptxAiConfig = { connection: { kind: 'endpoint', api: '/api/chat' } };

let container: HTMLDivElement;
let root: Root;
let content: Uint8Array;

beforeEach(async () => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	window.happyDOM?.setViewport({ width: 1600, height: 950 });
	lazyPanel.received = [];
	lazyPanel.mounts = 0;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	content = await handler.save(data.slides);
	handler.dispose();
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

const tick = () =>
	act(async () => {
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	});

const render = (ai: ViewerProps['ai'], ref?: React.Ref<ViewerHandle>) =>
	act(async () => {
		root.render(<PowerPointViewer ref={ref} content={content} canEdit ai={ai} />);
	});

/** Mount the viewer, wait for its deck, and open the assistant from the toolbar. */
async function mountOpen(ai: ViewerProps['ai']): Promise<void> {
	const ref = createRef<ViewerHandle>();
	await render(ai, ref);
	for (let attempt = 0; attempt < 80 && (ref.current?.getSlideCount() ?? 0) === 0; attempt += 1) {
		await tick();
	}
	expect(ref.current?.getSlideCount()).toBeGreaterThan(0);
	const toggle = container.querySelector<HTMLButtonElement>(
		`button[aria-label="${translationsEn['pptx.toolbar.toggleAiAssistant']}"]`,
	);
	expect(toggle).not.toBeNull();
	await act(async () => {
		toggle?.click();
	});
	expect(container.querySelector('[data-testid="lazy-panel"]')).not.toBeNull();
}

describe('powerPointViewer ai panel path', () => {
	it('gives the chat panel the handle config and the live viewer state', async () => {
		await mountOpen(createAiAssistant(config));
		const props = lazyPanel.received.at(-1);
		expect(props).toBeDefined();
		const { config: received, bridge, aiPanel, onClose } = props as AiChatPanelLazyProps;
		expect(received).toBe(config);
		expect(bridge.getSlides).toBeTypeOf('function');
		expect(aiPanel.isOpen).toBeTruthy();
		expect(onClose).toBe(aiPanel.close);
		await act(async () => {
			container.querySelector<HTMLButtonElement>('[data-testid="lazy-panel"]')?.click();
		});
		expect(container.querySelector('[data-testid="lazy-panel"]')).toBeNull();
	});

	it('gives the chat panel a plain config as it is', async () => {
		await mountOpen(config);
		expect(lazyPanel.received.at(-1)?.config).toBe(config);
	});

	it('treats a hand-built handle-shaped object as a plain config', async () => {
		const Panel = vi.fn(() => null);
		const handBuilt = { binding: 'react', config, panel: Panel };
		await mountOpen(handBuilt as unknown as PptxAiConfig);
		expect(Panel).not.toHaveBeenCalled();
		expect(lazyPanel.received.at(-1)?.config).toBe(handBuilt);
	});

	it('keeps the chat mounted when the host recreates the handle on every render', async () => {
		await mountOpen(createAiAssistant({ ...config }));
		const latest = { ...config };
		await render(createAiAssistant({ ...config }));
		await render(createAiAssistant(latest));
		expect(lazyPanel.mounts).toBe(1);
		expect(lazyPanel.received.at(-1)?.config).toBe(latest);
	});
});
