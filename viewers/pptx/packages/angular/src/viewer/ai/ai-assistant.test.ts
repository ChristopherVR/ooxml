/**
 * `pptx-angular-viewer/ai` (the `ai/` secondary entry point): only `createAiAssistant` makes
 * a handle (through the root's config-only `createAiAssistantHandle`), the handle carries its
 * config only, and the viewer template hands that config to the panel, whose `AiChatService`
 * loads the chat runtime itself. No panel input or service parameter takes a loader. A
 * hand-built object is a plain config, never a handle. Angular has no TestBed here (see
 * `vitest.config.ts`), so the template and input wiring is read from the component sources,
 * as the other `*.contract.test.ts` files do; `e2e/pptx/ai-panel-parity.spec.ts` opens the
 * panel through the handle in the Angular demo.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { DestroyRef, Injector, runInInjectionContext } from '@angular/core';
import type { PptxAiBridge, PptxAiConfig } from 'ooxml-ui/pptx/ai';
import { createAiAssistantRegistry } from 'ooxml-ui/pptx/ai/host';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createAiAssistant } from '../../../ai/public-api';
import { aiAssistants, createAiAssistantHandle } from '../ai-assistant-registry';
import { AiChatService } from './ai-chat.service';

const runtime = vi.hoisted(() => ({
	isAiAvailable: vi.fn(),
	createVanillaChat: vi.fn(),
}));
vi.mock(import('ooxml-ui/pptx/ai'), () => runtime as never);

const config: PptxAiConfig = { connection: { kind: 'endpoint', api: '/api/chat' } };
const bridge = {} as PptxAiBridge;
const injectors: ReturnType<typeof Injector.create>[] = [];

function service(): AiChatService {
	const injector = Injector.create({ providers: [] });
	injectors.push(injector);
	return runInInjectionContext(injector, () => new AiChatService());
}

function readSource(file: string): string {
	return readFileSync(path.join(import.meta.dirname, file), 'utf8');
}

beforeEach(() => {
	runtime.isAiAvailable.mockReset().mockResolvedValue(true);
	runtime.createVanillaChat.mockReset().mockResolvedValue({
		subscribe: () => () => undefined,
		getSnapshot: () => ({ messages: [], status: 'ready', error: undefined }),
		proposals: { list: () => [] },
	});
});

afterEach(() => {
	for (const injector of injectors.splice(0)) {
		if (!injector.get(DestroyRef).destroyed) {
			injector.destroy();
		}
	}
});

describe('createAiAssistant (angular)', () => {
	it('returns an opaque, frozen handle that carries its config only', () => {
		const handle = createAiAssistant(config);
		expect(Reflect.ownKeys(handle)).toStrictEqual([]);
		expect(Object.isFrozen(handle)).toBeTruthy();
		expect(aiAssistants.resolve(handle)).toStrictEqual({ config, panel: undefined });
	});

	it('takes nothing but a config through the root export it is built on', () => {
		const loader = vi.fn();
		const handle = (createAiAssistantHandle as (...args: unknown[]) => unknown)(config, loader);
		expect(aiAssistants.resolve(handle as PptxAiConfig)).toStrictEqual({
			config,
			panel: undefined,
		});
		expect(loader).not.toHaveBeenCalled();
	});

	it('reads a hand-built object, or a handle of a host-made registry, as a plain config', () => {
		const handBuilt = { binding: 'angular', config, panel: vi.fn() };
		expect(aiAssistants.resolve(handBuilt as unknown as PptxAiConfig)).toStrictEqual({
			config: handBuilt,
			panel: undefined,
		});
		expect(handBuilt.panel).not.toHaveBeenCalled();
		const hostMade = createAiAssistantRegistry<PptxAiConfig, () => unknown>({ owner: 'host' });
		const foreign = hostMade.register(config, vi.fn());
		expect(aiAssistants.resolve(foreign, { development: false })).toStrictEqual({
			config: foreign,
			panel: undefined,
		});
	});
});

describe('aiChatService', () => {
	it('starts the chat for a handle with its own runtime import', async () => {
		const resolved = aiAssistants.resolve(createAiAssistant(config));
		const chat = service();
		chat.init(bridge, resolved?.config ?? ({} as PptxAiConfig));
		await vi.waitFor(() => expect(chat.state()).toBe('ready'));
		expect(runtime.createVanillaChat).toHaveBeenCalledExactlyOnceWith({ bridge, config });
	});

	it('takes no runtime loader from a caller', async () => {
		expect(AiChatService.prototype.init).toHaveLength(2);
		const hostLoader = vi.fn();
		const chat = service();
		(chat.init as (...args: unknown[]) => void)(bridge, config, hostLoader);
		await vi.waitFor(() => expect(chat.state()).toBe('ready'));
		expect(hostLoader).not.toHaveBeenCalled();
		expect(runtime.createVanillaChat).toHaveBeenCalledExactlyOnceWith({ bridge, config });
	});
});

describe('chat panel inputs', () => {
	it('declares only the bridge, config and panel width inputs, and no loader', () => {
		const source = readSource('ai-chat-panel.component.ts');
		const inputs = [...source.matchAll(/readonly (\w+) = input\b/gu)].map(([, name]) => name);
		expect(inputs).toStrictEqual(['bridge', 'config', 'panelWidth']);
		expect(source).toContain('this.chat.init(this.bridge(), this.config());');
	});
});

describe('viewer template wiring', () => {
	const source = readSource(path.join('..', 'power-point-viewer.component.ts'));

	it("reads the ai input through the package's registry", () => {
		expect(source).toMatch(
			/aiAssistants\.resolve\(this\.ai\(\), \{ development: isDevMode\(\) \}\)/u,
		);
	});

	it('hands the resolved config, and nothing else of the handle, to the chat panel', () => {
		const panel = source.slice(source.indexOf('@if (aiAssistant(); as assistant)'));
		const element = panel.slice(0, panel.indexOf('/>'));
		expect(element).toMatch(/<pptx-ai-chat-panel[^>]*\[config\]="assistant\.config"/u);
		expect([...element.matchAll(/\[(\w+)\]=/gu)].map(([, name]) => name)).toStrictEqual([
			'bridge',
			'config',
		]);
	});
});
