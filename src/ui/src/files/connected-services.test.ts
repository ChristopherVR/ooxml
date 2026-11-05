// @vitest-environment happy-dom
import type { CloudAccount, CloudFileProvider } from 'ooxml-core/embedding';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { OfficeUiConnectedServices, defineConnectedServices } from './connected-services';

defineConnectedServices();

function provider(
	options: { signOut?: boolean; account?: CloudAccount | null; fail?: boolean } = {},
) {
	let account = options.account ?? null;
	const listeners = new Set<() => void>();
	const notify = () => {
		for (const listener of listeners) listener();
	};
	const value: CloudFileProvider = {
		id: 'onedrive',
		label: 'OneDrive',
		account: () => account,
		signIn: vi.fn(async () => {
			if (options.fail) throw new Error('popup closed');
			account = { name: 'Ada', email: 'ada@example.com' };
			notify();
			return account;
		}),
		subscribe: (listener) => {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
		list: async () => ({ entries: [] }),
		open: async () => {
			throw new Error('not used');
		},
	};
	if (options.signOut !== false)
		value.signOut = vi.fn(async () => {
			account = null;
			notify();
		});
	return { value, setAccount: (next: CloudAccount | null) => ((account = next), notify()) };
}

async function mount(providers: CloudFileProvider[]) {
	const element = document.createElement(
		'office-ui-connected-services',
	) as OfficeUiConnectedServices;
	element.providers = providers;
	document.body.append(element);
	await element.updateComplete;
	return element;
}

const root = (element: OfficeUiConnectedServices) => element.shadowRoot!;
const text = (element: OfficeUiConnectedServices) =>
	root(element).querySelector('.account')?.textContent?.trim();
const button = (element: OfficeUiConnectedServices) =>
	root(element).querySelector<HTMLButtonElement>('button');
const settle = async (element: OfficeUiConnectedServices) => {
	await new Promise((resolve) => setTimeout(resolve));
	await element.updateComplete;
};

afterEach(() => {
	document.body.replaceChildren();
});

describe('office-ui-connected-services', () => {
	it('renders nothing without providers', async () => {
		const element = await mount([]);
		expect(root(element).querySelector('section')).toBeNull();
	});

	it('connects and disconnects through the provider, and reports the change', async () => {
		const { value } = provider();
		const element = await mount([value]);
		const changes: unknown[] = [];
		element.addEventListener('office-connected-service-change', (event) =>
			changes.push((event as CustomEvent).detail),
		);
		expect(text(element)).toBe('Not connected');
		expect(button(element)?.textContent?.trim()).toBe('Connect');
		button(element)!.click();
		await settle(element);
		expect(value.signIn).toHaveBeenCalledOnce();
		expect(text(element)).toBe('Ada (ada@example.com)');
		expect(button(element)?.textContent?.trim()).toBe('Disconnect');
		button(element)!.click();
		await settle(element);
		expect(value.signOut).toHaveBeenCalledOnce();
		expect(text(element)).toBe('Not connected');
		expect(changes).toEqual([
			{ providerId: 'onedrive', connected: true },
			{ providerId: 'onedrive', connected: false },
		]);
	});

	it('follows sign-in elsewhere, shows a host-managed session and stops listening when removed', async () => {
		const { value, setAccount } = provider({ signOut: false });
		const element = await mount([value]);
		setAccount({ name: 'Grace' });
		await element.updateComplete;
		expect(text(element)).toBe('Grace');
		expect(button(element)).toBeNull();
		expect(root(element).querySelector('.managed')?.textContent).toBe(
			'Managed by your organization',
		);
		element.remove();
		const update = vi.spyOn(element, 'requestUpdate');
		setAccount(null);
		expect(update).not.toHaveBeenCalled();
	});

	it('reports a failed connect and accepts translated labels', async () => {
		const { value } = provider({ fail: true });
		const element = await mount([value]);
		element.labels = { connect: 'Verbinden', failed: '{provider} fehlgeschlagen' };
		await element.updateComplete;
		button(element)!.click();
		await settle(element);
		expect(root(element).querySelector('[role="alert"]')?.textContent).toBe(
			'OneDrive fehlgeschlagen',
		);
		expect(button(element)?.textContent?.trim()).toBe('Verbinden');
		expect(button(element)?.disabled).toBe(false);
	});
});
