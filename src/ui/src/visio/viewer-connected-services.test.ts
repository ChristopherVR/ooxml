// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import type { CloudFileProvider } from 'ooxml-core/embedding';
import { mountViewer } from './binding';

afterEach(() => document.body.replaceChildren());

const provider: CloudFileProvider = {
	id: 'cloud',
	label: 'Cloud',
	account: () => ({ name: 'Ada' }),
	signIn: async () => ({ name: 'Ada' }),
	list: async () => ({ entries: [] }),
	open: async () => {
		throw new Error('unused');
	},
};

function setup(options: Parameters<typeof mountViewer>[1] = {}) {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, options);
	const root = viewer.element.shadowRoot!;
	const list = () =>
		root.querySelector<HTMLElement & { shadowRoot: ShadowRoot }>(
			'[data-backstage-page="account"] office-ui-connected-services',
		)!;
	return { viewer, list };
}

const rows = (list: HTMLElement & { shadowRoot: ShadowRoot }) =>
	[...list.shadowRoot.querySelectorAll('.service')].map((row) => row.getAttribute('data-provider'));

describe('Visio Account > Connected services', () => {
	it('lists the file providers on the Account page', () => {
		const { viewer, list } = setup({ fileProviders: [provider] });
		expect(rows(list())).toEqual(['cloud']);
		expect(list().shadowRoot.textContent).toContain('Connected services');
		viewer.destroy();
	});

	it('follows fileProviders and shows nothing without providers', () => {
		const { viewer, list } = setup();
		expect(list().shadowRoot.querySelector('.service')).toBeNull();
		viewer.element.fileProviders = [provider];
		expect(rows(list())).toEqual(['cloud']);
		viewer.element.fileProviders = [];
		expect(list().shadowRoot.querySelector('.service')).toBeNull();
		viewer.destroy();
	});

	it('uses the display language', () => {
		const { viewer, list } = setup({ fileProviders: [provider] });
		viewer.element.setAttribute('locale', 'de');
		expect(list().shadowRoot.textContent).toContain('Verbundene Dienste');
		viewer.destroy();
	});
});
