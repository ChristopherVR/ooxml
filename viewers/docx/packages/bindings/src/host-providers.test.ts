// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { EDITOR_PROP_KEYS, mountEditor } from './index';

afterEach(() => document.body.replaceChildren());

const host = () => document.body.appendChild(document.createElement('div'));

describe('host providers through the shared contract', () => {
	const search = { name: 'Videos', search: async () => [] };
	const store = { read: () => undefined, write: () => {} };
	const signer = { certificate: () => new Uint8Array(), sign: async () => new Uint8Array() };

	it.each(['onlineVideoSearch', 'buildingBlocksStore', 'documentSigner'] as const)(
		'lists %s',
		(key) => expect(EDITOR_PROP_KEYS).toContain(key),
	);

	it('forwards each provider and clears it when left out', () => {
		const binding = mountEditor(host(), {
			onlineVideoSearch: search,
			buildingBlocksStore: store,
			documentSigner: signer,
		});
		expect(binding.element.onlineVideoSearch).toBe(search);
		expect(binding.element.buildingBlocksStore).toBe(store);
		expect(binding.element.documentSigner).toBe(signer);
		binding.update({});
		expect(binding.element.onlineVideoSearch).toBeUndefined();
		expect(binding.element.buildingBlocksStore).toBeUndefined();
		expect(binding.element.documentSigner).toBeUndefined();
		binding.destroy();
	});
});
