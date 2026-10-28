/**
 * `pptx-svelte-viewer/ai`: only `createAiAssistant` makes a handle, the handle carries the
 * dock's panel loader, and the `ai` prop opens the same panel for a handle and a plain config.
 * A hand-built object is a plain config, never a handle. The chat panel is replaced by a
 * recorder, so nothing loads `@ai-sdk/svelte` or the `ai` SDK.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createAiAssistant } from '../ai';
import type { PptxAiAssistant } from '../ai';
import { aiAssistants } from './ai-assistant-registry';
import PowerPointViewer from './PowerPointViewer.svelte';

const recorded = vi.hoisted(() => ({ configs: [] as unknown[] }));

vi.mock(import('./components/ai/AiChatPanel.svelte'), () => ({
	// A stand-in component: Svelte calls it with an anchor node and its props.
	default: ((anchor: Comment, props: { config: unknown }) => {
		recorded.configs.push(props.config);
		const node = document.createElement('div');
		node.setAttribute('data-testid', 'recorded-ai-panel');
		anchor.before(node);
		return {};
	}) as never,
}));

const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');
const config: PptxAiConfig = { connection: { kind: 'endpoint', api: '/api/chat' } };

let cleanup: (() => void) | undefined;

beforeEach(() => {
	recorded.configs = [];
});

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

/** Mount the viewer with `ai`, open the assistant from the ribbon, wait for the panel. */
async function openWith(ai: PptxAiConfig | PptxAiAssistant): Promise<void> {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const onload = vi.fn();
	const instance = mount(PowerPointViewer, {
		target,
		props: { source: new Uint8Array(readFileSync(FIXTURE)), editable: true, onload, ai },
	});
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	flushSync();
	await vi.waitFor(() => expect(onload).toHaveBeenCalledOnce(), { timeout: 15000 });
	flushSync();
	const toggle = target.querySelector<HTMLButtonElement>(
		'button[aria-label="Toggle AI assistant"]',
	);
	expect(toggle).not.toBeNull();
	toggle?.click();
	flushSync();
	await vi.waitFor(() => expect(recorded.configs).toHaveLength(1));
	expect(target.querySelector('[data-testid="recorded-ai-panel"]')).not.toBeNull();
}

describe('createAiAssistant (svelte)', () => {
	it('returns an opaque, frozen handle that carries the panel loader', async () => {
		const handle = createAiAssistant(config);
		expect(Reflect.ownKeys(handle)).toStrictEqual([]);
		expect(Object.isFrozen(handle)).toBeTruthy();
		const resolved = aiAssistants.resolve(handle);
		expect(resolved?.config).toBe(config);
		const panelModule = await import('./components/ai/AiChatPanel.svelte');
		await expect(resolved?.panel?.()).resolves.toBe(panelModule);
	});

	it('reads a hand-built object, or a handle of a host-made registry, as a plain config', () => {
		const handBuilt = { binding: 'svelte', config, panel: vi.fn() };
		expect(aiAssistants.resolve(handBuilt as unknown as PptxAiConfig)?.panel).toBeUndefined();
		const hostMade = createAiAssistantRegistry<PptxAiConfig, () => Promise<never>>({
			owner: 'host',
		});
		const foreign = hostMade.register(config, vi.fn());
		expect(aiAssistants.resolve(foreign, { development: false })?.panel).toBeUndefined();
	});
});

describe('the ai prop', () => {
	it('opens the chat panel of a createAiAssistant handle with its config', async () => {
		await openWith(createAiAssistant(config));
		expect(recorded.configs[0]).toBe(config);
	});

	it('loads the panel through the loader the handle carries', async () => {
		const loader = vi.fn(() => import('./components/ai/AiChatPanel.svelte'));
		// What `createAiAssistant` does, through the package's private registry, with a loader
		// the test can observe.
		await openWith(aiAssistants.register(config, loader));
		expect(loader).toHaveBeenCalledOnce();
	});

	it('opens the same panel for a plain config', async () => {
		await openWith(config);
		expect(recorded.configs[0]).toBe(config);
	});

	it('passes a hand-built handle-shaped object on as a plain config', async () => {
		const panel = vi.fn();
		const handBuilt = { binding: 'svelte', config, panel };
		await openWith(handBuilt as unknown as PptxAiConfig);
		expect(recorded.configs[0]).toBe(handBuilt);
		expect(panel).not.toHaveBeenCalled();
	});
});
