// @vitest-environment happy-dom
/**
 * The `ai` prop end to end through the real `PowerPointViewer`: a `createAiAssistant` handle
 * (`pptx-react-viewer/ai`) and a plain `PptxAiConfig` both open the chat panel from the
 * toolbar toggle, and a hand-built object is read as a plain config, never as a handle.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createAiAssistant } from '../ai';

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

const tick = () =>
	act(async () => {
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	});

async function mount(props: Partial<Omit<ViewerProps, 'content'>>): Promise<ViewerHandle> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	const content = await handler.save(data.slides);
	handler.dispose();
	const ref = createRef<ViewerHandle>();
	await act(async () => {
		root.render(<PowerPointViewer ref={ref} content={content} canEdit {...props} />);
	});
	for (let attempt = 0; attempt < 80; attempt += 1) {
		if ((ref.current?.getSlideCount() ?? 0) > 0) {
			await tick();
			return ref.current as ViewerHandle;
		}
		await tick();
	}
	throw new Error('the viewer never finished loading its deck');
}

const aiToggle = () =>
	container.querySelector<HTMLButtonElement>(
		`button[aria-label="${translationsEn['pptx.toolbar.toggleAiAssistant']}"]`,
	);
/** The viewer's chat panel (`AiChatPanel`, behind `AiChatPanelLazy`). */
const chatPanel = () => container.querySelector('[data-pptx-ai-panel]');
/** The Suspense fallback of `AiChatPanelLazy` while the panel chunk loads. */
const lazyFallback = () => container.querySelector('.border-l .animate-spin');

async function openChatPanel(): Promise<Element> {
	const toggle = aiToggle();
	expect(toggle).not.toBeNull();
	await act(async () => {
		toggle?.click();
	});
	// The lazy chunk may already be loaded by an earlier file in this worker.
	expect(lazyFallback() ?? chatPanel()).not.toBeNull();
	for (let attempt = 0; attempt < 100 && !chatPanel(); attempt += 1) {
		await tick();
	}
	const panel = chatPanel();
	expect(panel).not.toBeNull();
	return panel as Element;
}

describe('powerPointViewer ai prop', () => {
	it('opens and closes the chat panel of a createAiAssistant handle', async () => {
		await mount({ ai: createAiAssistant(config) });
		expect(chatPanel()).toBeNull();
		const panel = await openChatPanel();
		const close = panel.querySelector<HTMLButtonElement>(
			`button[aria-label="${translationsEn['pptx.ai.close']}"]`,
		);
		expect(close).not.toBeNull();
		await act(async () => {
			close?.click();
		});
		expect(chatPanel()).toBeNull();
	});

	it('shows no assistant for a handle when the ai feature is turned off', async () => {
		await mount({
			ai: createAiAssistant(config),
			customization: { disabledFeatures: ['ai'] },
		});
		expect(aiToggle()).toBeNull();
		expect(chatPanel()).toBeNull();
	});

	it('keeps a plain config on the built-in lazy panel', async () => {
		await mount({ ai: config });
		await openChatPanel();
	});

	it('reads a hand-built handle-shaped object as a plain config', async () => {
		const Panel = vi.fn(() => null);
		const forged = { binding: 'react', config, panel: Panel };
		await mount({ ai: forged as unknown as PptxAiConfig });
		const panel = await openChatPanel();
		// It has no `connection`, so the session cannot start, as for any unusable config.
		for (
			let attempt = 0;
			attempt < 100 && !panel.textContent?.includes(translationsEn['pptx.ai.unavailableTitle']);
			attempt += 1
		) {
			await tick();
		}
		expect(panel.textContent).toContain(translationsEn['pptx.ai.unavailableTitle']);
		expect(Panel).not.toHaveBeenCalled();
	});
});
