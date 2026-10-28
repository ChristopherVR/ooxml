/**
 * The React `createAiAssistant` handle is opaque: only the factory makes one, it carries the
 * viewer's lazy chat panel, and no other value, whatever its shape, is read as a handle.
 */
import type { PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';
import { describe, expect, it } from 'vitest';

import { createAiAssistant } from '../ai';
import { aiAssistants } from './ai-assistant-registry';
import { AiChatPanelLazy } from './components/ai/AiChatPanelLazy';

const config: PptxAiConfig = { connection: { kind: 'endpoint', api: '/api/chat' } };

describe('createAiAssistant', () => {
	it('returns an opaque, frozen handle with no readable members', () => {
		const assistant = createAiAssistant(config);
		expect(Reflect.ownKeys(assistant)).toStrictEqual([]);
		expect(Object.isFrozen(assistant)).toBeTruthy();
		expect(aiAssistants.resolve(assistant)).toStrictEqual({ config, panel: AiChatPanelLazy });
	});

	it('gives every handle its own config and the same chat panel', () => {
		const other: PptxAiConfig = { ...config, writePolicy: 'auto' };
		const first = createAiAssistant(config);
		const second = createAiAssistant(other);
		expect(second).not.toBe(first);
		expect(aiAssistants.resolve(first)).toStrictEqual({ config, panel: AiChatPanelLazy });
		expect(aiAssistants.resolve(second)).toStrictEqual({
			config: other,
			panel: AiChatPanelLazy,
		});
	});

	it('reads no other value as a handle', () => {
		const assistant = createAiAssistant(config);
		const handBuilt = { binding: 'react', config, panel: () => null };
		// A registry a host creates from the public factory is not this package's registry.
		const hostMade = createAiAssistantRegistry<PptxAiConfig, () => null>({ owner: 'host' });
		const foreign = hostMade.register(config, () => null);
		for (const value of [config, {}, Object.freeze({}), { ...assistant }, handBuilt, foreign]) {
			expect(aiAssistants.resolve(value as PptxAiConfig, { development: false })).toStrictEqual({
				config: value,
				panel: undefined,
			});
		}
	});
});
