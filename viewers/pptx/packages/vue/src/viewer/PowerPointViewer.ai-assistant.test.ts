/**
 * `pptx-vue-viewer/ai`: only `createAiAssistant` makes a handle, the handle carries the
 * viewer's async chat panel, and the `ai` prop opens the same panel for a handle and a plain
 * config. A hand-built object is a plain config, never a handle. The chat panel is replaced
 * by a recorder, so nothing loads `@ai-sdk/vue` or the `ai` SDK.
 */
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { defineComponent, h } from 'vue';
import type { Component } from 'vue';

import { createAiAssistant } from '../ai';
import type { PptxAiAssistant } from '../ai';
import { aiAssistants } from './ai-assistant-registry';
import { AiChatPanelLazy } from './components/ai';
import PowerPointViewer from './PowerPointViewer.vue';

const recorded = vi.hoisted(() => ({ configs: [] as unknown[] }));

vi.mock(import('./components/ai/AiChatPanel.vue'), async () => {
	const vue = await import('vue');
	return {
		// `defineAsyncComponent` unwraps a module by its `__esModule` flag.
		__esModule: true,
		default: vue.defineComponent({
			props: { config: { type: Object, required: true } },
			setup(props) {
				recorded.configs.push(props.config);
				return () => vue.h('div', { 'data-testid': 'recorded-ai-panel' });
			},
		}) as never,
	};
});

const config: PptxAiConfig = { connection: { kind: 'endpoint', api: '/api/chat' } };

let wrapper: VueWrapper | undefined;

beforeEach(() => {
	recorded.configs = [];
});

afterEach(() => {
	wrapper?.unmount();
	wrapper = undefined;
});

/** Mount the viewer with `ai`, open the assistant from the ribbon, wait for the panel. */
async function openWith(ai: PptxAiConfig | PptxAiAssistant): Promise<VueWrapper> {
	wrapper = mount(PowerPointViewer, { props: { canEdit: true, ai }, attachTo: document.body });
	await flushPromises();
	const toggle = wrapper
		.findAll('button')
		.find((button) => button.attributes('aria-label') === 'Toggle AI assistant');
	expect(toggle).toBeDefined();
	await toggle?.trigger('click');
	await vi.waitFor(async () => {
		await flushPromises();
		expect(recorded.configs).toHaveLength(1);
	});
	expect(wrapper.find('[data-testid="recorded-ai-panel"]').exists()).toBeTruthy();
	return wrapper;
}

describe('createAiAssistant (vue)', () => {
	it('returns an opaque, frozen handle that carries the async chat panel', () => {
		const handle = createAiAssistant(config);
		expect(Reflect.ownKeys(handle)).toStrictEqual([]);
		expect(Object.isFrozen(handle)).toBeTruthy();
		expect(aiAssistants.resolve(handle)).toStrictEqual({ config, panel: AiChatPanelLazy });
	});

	it('reads a hand-built object, or a handle of a host-made registry, as a plain config', () => {
		const handBuilt = { binding: 'vue', config, panel: AiChatPanelLazy };
		expect(aiAssistants.resolve(handBuilt as unknown as PptxAiConfig)?.panel).toBeUndefined();
		const hostMade = createAiAssistantRegistry<PptxAiConfig, Component>({ owner: 'host' });
		const foreign = hostMade.register(config, AiChatPanelLazy);
		expect(aiAssistants.resolve(foreign, { development: false })?.panel).toBeUndefined();
	});
});

describe('the ai prop', () => {
	it('opens the chat panel of a createAiAssistant handle with its config', async () => {
		await openWith(createAiAssistant(config));
		expect(recorded.configs[0]).toBe(config);
	});

	it('opens the same panel for a plain config', async () => {
		await openWith(config);
		expect(recorded.configs[0]).toStrictEqual(config);
	});

	it('passes a hand-built handle-shaped object on as a plain config', async () => {
		const panel = defineComponent({
			setup: () => () => h('div', { 'data-testid': 'hand-built-panel' }),
		});
		const handBuilt = { binding: 'vue', config, panel };
		const mounted = await openWith(handBuilt as unknown as PptxAiConfig);
		expect(recorded.configs[0]).toStrictEqual(handBuilt);
		expect(mounted.find('[data-testid="hand-built-panel"]').exists()).toBeFalsy();
	});
});
