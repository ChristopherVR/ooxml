/**
 * `pptx-vanilla-viewer/ai`: only `createAiAssistant` makes a handle, the handle carries the
 * viewer's panel loader, and the `ai` option opens the same panel for a handle and a plain
 * config. A hand-built object is a plain config, never a handle. The panel module is replaced
 * by a recorder, so nothing loads the `ai` SDK.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createAiAssistant } from './ai';
import type { PptxAiAssistant } from './ai';
import { aiAssistants } from './viewer/ai-assistant-registry';
import { createPptxViewer } from './viewer/PptxViewer';
import type { PptxViewerInstance } from './viewer/types';

const panelModule = vi.hoisted(() => ({
	created: [] as { config: unknown }[],
}));

vi.mock(import('./viewer/ai/ai-panel'), () => ({
	createAiPanel: vi.fn(async (deps: { config: unknown }) => {
		panelModule.created.push(deps);
		return { destroy: () => undefined };
	}),
}));

const config: PptxAiConfig = { connection: { kind: 'endpoint', api: '/api/chat' } };

let active: PptxViewerInstance[] = [];

beforeEach(() => {
	panelModule.created = [];
});

afterEach(() => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
});

/** Mount a viewer with `ai`, open the assistant from its toggle, and wait for the panel. */
async function openWith(ai: PptxAiConfig | PptxAiAssistant): Promise<HTMLElement> {
	const container = document.createElement('div');
	document.body.appendChild(container);
	active.push(createPptxViewer(container, { ai }));
	const toggle = container.querySelector<HTMLButtonElement>('.pptxv-ai-toggle');
	expect(toggle).not.toBeNull();
	toggle?.click();
	await vi.waitFor(() => expect(panelModule.created).toHaveLength(1));
	const host = container.querySelector<HTMLElement>('[data-pptx-ai-panel]');
	expect(host?.hidden).toBeFalsy();
	return container;
}

describe('createAiAssistant (vanilla)', () => {
	it('returns an opaque, frozen handle that carries the panel loader', async () => {
		const handle = createAiAssistant(config);
		expect(Reflect.ownKeys(handle)).toStrictEqual([]);
		expect(Object.isFrozen(handle)).toBeTruthy();
		const resolved = aiAssistants.resolve(handle);
		expect(resolved?.config).toBe(config);
		const loaded = await resolved?.panel?.();
		expect(loaded?.createAiPanel).toBe((await import('./viewer/ai/ai-panel')).createAiPanel);
	});

	it('reads a hand-built object, or a handle of a host-made registry, as a plain config', () => {
		const handBuilt = { binding: 'vanilla', config, panel: vi.fn() };
		expect(aiAssistants.resolve(handBuilt as unknown as PptxAiConfig)).toStrictEqual({
			config: handBuilt,
			panel: undefined,
		});
		const hostMade = createAiAssistantRegistry<PptxAiConfig, () => Promise<never>>({
			owner: 'host',
		});
		const foreign = hostMade.register(config, vi.fn());
		expect(aiAssistants.resolve(foreign, { development: false })?.panel).toBeUndefined();
	});
});

describe('the ai option', () => {
	it('opens the panel of a createAiAssistant handle with its config', async () => {
		await openWith(createAiAssistant(config));
		expect(panelModule.created[0]?.config).toBe(config);
	});

	it('loads the panel through the loader the handle carries', async () => {
		const loader = vi.fn(() => import('./viewer/ai/ai-panel'));
		// What `createAiAssistant` does, through the package's private registry, with a loader
		// the test can observe.
		await openWith(aiAssistants.register(config, loader));
		expect(loader).toHaveBeenCalledOnce();
		expect(panelModule.created[0]?.config).toBe(config);
	});

	it('opens the same panel for a plain config', async () => {
		await openWith(config);
		expect(panelModule.created[0]?.config).toBe(config);
	});

	it.each([false, '', 0, undefined, null])(
		'shows no toggle and opens no panel for a falsy ai (%j)',
		async (ai) => {
			const container = document.createElement('div');
			document.body.appendChild(container);
			active.push(createPptxViewer(container, { ai: ai as unknown as PptxAiConfig }));
			expect(container.querySelector('.pptxv-ai-toggle')).toBeNull();
			expect(container.querySelector('[data-pptx-ai-panel]')).toBeNull();
			await new Promise<void>((resolve) => {
				setTimeout(resolve, 0);
			});
			expect(panelModule.created).toStrictEqual([]);
		},
	);

	it('passes a hand-built handle-shaped object on as a plain config', async () => {
		const panel = vi.fn();
		const handBuilt = { binding: 'vanilla', config, panel };
		await openWith(handBuilt as unknown as PptxAiConfig);
		expect(panelModule.created[0]?.config).toBe(handBuilt);
		expect(panel).not.toHaveBeenCalled();
	});
});
