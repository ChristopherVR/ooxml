/**
 * The `createAiAssistant` handle registry factory: a registry reads only the handles it
 * created, every other value (including a handle from another registry) is a plain config,
 * and a handle a registry cannot read is reported once, in development only.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createAiAssistantRegistry } from './ai-assistant';
import type { PptxAiAssistant } from './ai-assistant';

interface Config {
	readonly connection: { readonly kind: 'endpoint'; readonly api: string };
	readonly writePolicy?: 'auto';
}
type Panel = () => Promise<unknown>;

const config: Config = { connection: { kind: 'endpoint', api: '/api/chat' } };
const panel: Panel = () => Promise.resolve({});

const registryOf = (owner: string) => createAiAssistantRegistry<Config, Panel>({ owner });

/** A second instance of this module, as a second copy of the package would load. */
async function secondCopy(): Promise<typeof import('./ai-assistant')> {
	const specifier = './ai-assistant.ts?second-copy';
	return (await import(/* @vite-ignore */ specifier)) as typeof import('./ai-assistant');
}

afterEach(() => {
	vi.unstubAllEnvs();
});

describe('createAiAssistantRegistry', () => {
	it('returns a frozen registry with only register and resolve', () => {
		const registry = registryOf('pptx-vue-viewer');
		expect(Object.keys(registry).sort()).toStrictEqual(['register', 'resolve']);
		expect(Object.isFrozen(registry)).toBeTruthy();
	});

	it('creates opaque, frozen handles with no readable members', () => {
		const registry = registryOf('pptx-vue-viewer');
		const handle = registry.register(config, panel);
		expect(Reflect.ownKeys(handle)).toStrictEqual([]);
		expect(Object.isFrozen(handle)).toBeTruthy();
		expect(Object.prototype.toString.call(handle)).toBe('[object PptxAiAssistant]');
		expect(registry.resolve(handle)).toStrictEqual({ config, panel });
	});

	it('gives every handle its own config', () => {
		const registry = registryOf('pptx-svelte-viewer');
		const other: Config = { ...config, writePolicy: 'auto' };
		const first = registry.register(config, panel);
		const second = registry.register(other, panel);
		expect(second).not.toBe(first);
		expect(registry.resolve(first)).toStrictEqual({ config, panel });
		expect(registry.resolve(second)).toStrictEqual({ config: other, panel });
	});
});

describe('resolve', () => {
	it('gives nothing for a missing value', () => {
		const registry = registryOf('pptx-react-viewer');
		expect(registry.resolve(undefined)).toBeUndefined();
		expect(registry.resolve(null)).toBeUndefined();
	});

	it('reads every value it did not create as a plain config, silently', () => {
		const registry = registryOf('pptx-vanilla-viewer');
		const warn = vi.fn<(message: string) => void>();
		const handle = registry.register(config, panel);
		const handBuilt = { config, panel };
		for (const value of [config, {}, Object.freeze({}), { ...handle }, handBuilt]) {
			expect(registry.resolve(value as Config, { development: true, warn })).toStrictEqual({
				config: value,
				panel: undefined,
			});
		}
		expect(warn).not.toHaveBeenCalled();
	});

	it('does not read a handle from another registry, and warns about it once', () => {
		const react = registryOf('pptx-react-viewer');
		// A registry a host creates itself is just another registry: no viewer reads it.
		const hostMade = registryOf('host');
		const warn = vi.fn<(message: string) => void>();
		const options = { development: true, warn };
		const foreign = hostMade.register(config, panel);
		expect(react.resolve(foreign, options)).toStrictEqual({ config: foreign, panel: undefined });
		react.resolve(foreign, options);
		expect(warn).toHaveBeenCalledOnce();
		expect(warn.mock.calls[0]?.[0]).toContain(
			'handle that pptx-react-viewer/ai did not create: it comes from another binding',
		);
	});

	it('warns once about a handle from another copy of the package', async () => {
		const copy = await secondCopy();
		expect(copy.createAiAssistantRegistry).not.toBe(createAiAssistantRegistry);
		const angular = registryOf('pptx-angular-viewer');
		const foreign = copy
			.createAiAssistantRegistry<Config, Panel>({ owner: 'pptx-angular-viewer' })
			.register(config, panel);
		const warn = vi.fn<(message: string) => void>();
		const options = { development: true, warn };
		expect(angular.resolve(foreign, options)?.panel).toBeUndefined();
		angular.resolve(foreign, options);
		expect(warn).toHaveBeenCalledOnce();
		expect(warn.mock.calls[0]?.[0]).toContain('another copy of the package');
		// An object created from a handle carries its tag too, but is still no handle.
		angular.resolve(Object.create(foreign) as PptxAiAssistant, options);
		expect(warn).toHaveBeenCalledTimes(2);
	});

	it('follows NODE_ENV by default and stays silent in production', () => {
		const react = registryOf('pptx-react-viewer');
		const other = registryOf('pptx-vue-viewer');
		const spy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		try {
			vi.stubEnv('NODE_ENV', 'production');
			react.resolve(other.register(config, panel));
			expect(spy).not.toHaveBeenCalled();
			vi.stubEnv('NODE_ENV', 'development');
			react.resolve(other.register(config, panel));
			expect(spy).toHaveBeenCalledOnce();
		} finally {
			spy.mockRestore();
		}
	});

	it('stays silent when the binding turns development checks off', () => {
		const warn = vi.fn<(message: string) => void>();
		const foreign = registryOf('pptx-vue-viewer').register(config, panel);
		registryOf('pptx-angular-viewer').resolve(foreign, { development: false, warn });
		expect(warn).not.toHaveBeenCalled();
	});
});
