import { DestroyRef, Injector, runInInjectionContext } from '@angular/core';
import type { PptxAiBridge, PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { AiChatService } from './ai-chat.service';

const runtime = vi.hoisted(() => ({
	loaded: vi.fn(),
	isAiAvailable: vi.fn(),
	createVanillaChat: vi.fn(),
}));
vi.mock(import('ooxml-ui/pptx/ai'), () => {
	runtime.loaded();
	return runtime;
});

const bridge = {} as PptxAiBridge;
const config = {} as PptxAiConfig;

const injectors: { injector: ReturnType<typeof Injector.create>; destroyRef: DestroyRef }[] = [];

function setup() {
	const injector = Injector.create({ providers: [] });
	injectors.push({ injector, destroyRef: injector.get(DestroyRef) });
	const service = runInInjectionContext(injector, () => new AiChatService());
	return { service, injector };
}

describe('aiChatService lazy runtime', () => {
	beforeEach(() => {
		runtime.isAiAvailable.mockReset().mockResolvedValue(true);
		runtime.createVanillaChat.mockReset();
	});

	afterEach(() => {
		for (const { injector, destroyRef } of injectors.splice(0)) {
			if (!destroyRef.destroyed) {
				injector.destroy();
			}
		}
	});

	it('does not load the runtime when the service is imported or constructed', () => {
		setup();
		expect(runtime.loaded).not.toHaveBeenCalled();
	});

	it('starts once and mirrors the lazily created controller', async () => {
		const unsubscribe = vi.fn();
		const snapshot = { messages: [], status: 'ready', error: undefined };
		const controller = {
			subscribe: vi.fn(() => unsubscribe),
			getSnapshot: () => snapshot,
			proposals: { list: () => [] },
			sendMessage: vi.fn(),
		};
		runtime.createVanillaChat.mockResolvedValue(controller);
		const { service, injector } = setup();
		service.init(bridge, config);
		service.init(bridge, config);
		await vi.waitFor(() => expect(service.state()).toBe('ready'));
		expect(runtime.createVanillaChat).toHaveBeenCalledExactlyOnceWith({ bridge, config });
		service.send(' hello ');
		expect(controller.sendMessage).toHaveBeenCalledWith('hello');
		injector.destroy();
		expect(unsubscribe).toHaveBeenCalledOnce();
	});

	it('reports a missing optional SDK', async () => {
		runtime.isAiAvailable.mockResolvedValue(false);
		const { service } = setup();
		service.init(bridge, config);
		await vi.waitFor(() => expect(service.state()).toBe('unavailable'));
		expect(runtime.createVanillaChat).not.toHaveBeenCalled();
	});

	it('reports runtime initialization failures', async () => {
		runtime.createVanillaChat.mockRejectedValue(new Error('Failed to load chat'));
		const { service } = setup();
		service.init(bridge, config);
		await vi.waitFor(() => expect(service.state()).toBe('error'));
		expect(service.initError()?.message).toBe('Failed to load chat');
	});

	it('does not start a session after the panel closes during loading', async () => {
		const { service, injector } = setup();
		service.init(bridge, config);
		injector.destroy();
		await new Promise((resolve) => {
			setTimeout(resolve, 0);
		});
		expect(runtime.createVanillaChat).not.toHaveBeenCalled();
	});
});
