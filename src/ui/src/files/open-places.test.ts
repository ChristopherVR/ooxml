// @vitest-environment happy-dom
import type { CloudFileProvider } from 'ooxml-core/embedding';
import { afterEach, describe, expect, it } from 'vitest';

import { DEVICE_PLACE, OfficeUiOpenPlaces, defineOpenPlaces } from './open-places';

defineOpenPlaces();

const provider = (id = 'onedrive', label = 'OneDrive'): CloudFileProvider => ({
	id,
	label,
	account: () => null,
	signIn: async () => ({ name: 'Ada' }),
	list: async () => ({ entries: [] }),
	open: async () => {
		throw new Error('not used');
	},
});

async function mount(providers: CloudFileProvider[]) {
	const element = document.createElement('office-ui-open-places') as OfficeUiOpenPlaces;
	element.providers = providers;
	element.innerHTML = '<button id="browse">Browse</button>';
	document.body.append(element);
	await element.updateComplete;
	return element;
}

const tabs = (element: OfficeUiOpenPlaces) => [
	...element.shadowRoot!.querySelectorAll<HTMLButtonElement>('[role="tab"]'),
];

afterEach(() => {
	document.body.replaceChildren();
});

describe('office-ui-open-places', () => {
	it('shows only the device content when there is no provider', async () => {
		const element = await mount([]);
		expect(tabs(element)).toHaveLength(0);
		expect(element.shadowRoot!.querySelector('office-ui-cloud-files')).toBeNull();
		expect(element.shadowRoot!.querySelector('slot')).not.toBeNull();
	});

	it('offers This device and one tab per provider, device first', async () => {
		const element = await mount([provider(), provider('box', 'Box')]);
		expect(tabs(element).map((tab) => tab.textContent)).toEqual(['This device', 'OneDrive', 'Box']);
		expect(tabs(element)[0]!.getAttribute('aria-selected')).toBe('true');
		expect(element.place).toBe(DEVICE_PLACE);
	});

	it('switches to a provider, shows its browser and reports the change', async () => {
		const element = await mount([provider()]);
		const places: string[] = [];
		element.addEventListener('office-open-place-change', (event) =>
			places.push((event as CustomEvent<{ place: string }>).detail.place),
		);
		tabs(element)[1]!.click();
		await element.updateComplete;
		const browser = element.shadowRoot!.querySelector('office-ui-cloud-files') as unknown as {
			providers: readonly CloudFileProvider[];
		};
		expect(browser.providers.map((p) => p.id)).toEqual(['onedrive']);
		expect(element.shadowRoot!.querySelector('slot')!.hidden).toBe(true);
		expect(places).toEqual(['onedrive']);
	});

	it('moves between places with the arrow keys and labels from the host', async () => {
		const element = await mount([provider()]);
		element.labels = { thisDevice: 'Dieses Gerät', places: 'Orte' };
		await element.updateComplete;
		expect(tabs(element)[0]!.textContent).toBe('Dieses Gerät');
		expect(element.shadowRoot!.querySelector('[role="tablist"]')!.getAttribute('aria-label')).toBe(
			'Orte',
		);
		tabs(element)[0]!.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }),
		);
		await element.updateComplete;
		expect(element.place).toBe('onedrive');
		tabs(element)[1]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
		await element.updateComplete;
		expect(element.place).toBe(DEVICE_PLACE);
	});

	it('falls back to This device when the chosen provider goes away', async () => {
		const element = await mount([provider()]);
		element.selected = 'onedrive';
		element.providers = [];
		await element.updateComplete;
		expect(element.place).toBe(DEVICE_PLACE);
	});
});
