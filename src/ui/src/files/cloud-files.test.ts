// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type {
	CloudAccount,
	CloudEntry,
	CloudFileProvider,
	CloudListing,
	CloudOpenOptions,
	CloudOpenedFile,
} from 'ooxml-core/embedding';
import { registerOfficeUi } from '../index';
import type { OfficeUiCloudFiles } from './cloud-files';

beforeAll(() => registerOfficeUi());
afterEach(() => {
	document.body.replaceChildren();
	vi.useRealTimers();
});

const folder = (id: string, name: string): CloudEntry => ({ id, name, type: 'folder' });
const file = (id: string, name: string, kind?: CloudEntry['kind']): CloudEntry => ({
	id,
	name,
	type: 'file',
	...(kind ? { kind } : {}),
	size: 2048,
	modified: Date.UTC(2026, 0, 5),
});

function opened(entry: CloudEntry): CloudOpenedFile {
	return {
		entry,
		kind: entry.kind ?? 'docx',
		fileName: entry.name,
		bytes: new Uint8Array([1]),
		canEdit: true,
		storage: { load: async () => null, save: async () => {} } as never,
	};
}

interface Fake extends CloudFileProvider {
	signedIn: boolean;
	notify(): void;
	lists: { folder: string | undefined; page: string | undefined; signal?: AbortSignal }[];
	searches: string[];
}

function fakeProvider(overrides: Partial<CloudFileProvider> = {}): Fake {
	const listeners = new Set<() => void>();
	const tree: Record<string, CloudListing> = {
		root: {
			entries: [folder('f1', 'Reports'), file('a', 'Letter.docx', 'docx'), file('t', 'x.txt')],
		},
		f1: { entries: [file('b', 'Q1.xlsx', 'xlsx')] },
	};
	const provider: Fake = {
		id: 'fake',
		label: 'Fake Drive',
		signedIn: false,
		lists: [],
		searches: [],
		account: () => (provider.signedIn ? ({ name: 'Ada' } as CloudAccount) : null),
		signIn: async () => {
			provider.signedIn = true;
			provider.notify();
			return { name: 'Ada' };
		},
		subscribe: (listener) => {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
		notify: () => listeners.forEach((listener) => listener()),
		list: async (id, options) => {
			provider.lists.push({
				folder: id,
				page: options?.page,
				signal: options?.signal as AbortSignal,
			});
			return tree[id ?? 'root'] ?? { entries: [] };
		},
		search: async (query) => {
			provider.searches.push(query);
			return { entries: [{ ...file('s', 'Found.docx', 'docx'), location: 'My files > Old' }] };
		},
		open: async (entry) => opened(entry),
		...overrides,
	};
	return provider;
}

async function mount(provider: CloudFileProvider, props: Partial<OfficeUiCloudFiles> = {}) {
	const element = document.createElement('office-ui-cloud-files') as OfficeUiCloudFiles;
	Object.assign(element, props);
	element.providers = [provider];
	document.body.append(element);
	await settle();
	return element;
}

const settle = async () => {
	for (let i = 0; i < 5; i++) await Promise.resolve();
};

const rows = (element: HTMLElement) => [
	...element.shadowRoot!.querySelectorAll<HTMLElement>('.row'),
];
const names = (element: HTMLElement) =>
	rows(element).map((row) => row.querySelector('b')!.textContent);
const button = (element: HTMLElement, selector: string) =>
	element.shadowRoot!.querySelector<HTMLButtonElement>(selector)!;

describe('office-ui-cloud-files', () => {
	it('signs in from the signed-out state and then lists files', async () => {
		const provider = fakeProvider();
		const element = await mount(provider);
		expect(element.shadowRoot!.textContent).toContain('Sign in to Fake Drive');
		button(element, '.primary').click();
		await settle();
		expect(names(element)).toEqual(['Reports', 'Letter.docx', 'x.txt']);
		expect(element.shadowRoot!.textContent).toContain('Ada');
	});

	it('shows the error when sign in rejects', async () => {
		const provider = fakeProvider({
			signIn: async () => {
				throw new Error('Popup blocked');
			},
		});
		const element = await mount(provider);
		button(element, '.primary').click();
		await settle();
		expect(element.shadowRoot!.querySelector('[role="alert"]')!.textContent).toBe('Popup blocked');
	});

	it('navigates into a folder, shows the breadcrumb and returns through it', async () => {
		const provider = fakeProvider();
		provider.signedIn = true;
		const element = await mount(provider);
		rows(element)[0]!.click();
		await settle();
		expect(names(element)).toEqual(['Q1.xlsx']);
		const crumbs = [...element.shadowRoot!.querySelectorAll('nav button')].map(
			(b) => b.textContent,
		);
		expect(crumbs).toEqual(['My files', 'Reports']);
		button(element, '[data-crumb="0"]').click();
		await settle();
		expect(names(element)).toHaveLength(3);
	});

	it('loads the next page and appends it', async () => {
		const provider = fakeProvider({
			list: async (_id, options) =>
				options?.page === 'p2'
					? { entries: [file('c', 'Two.docx', 'docx')] }
					: { entries: [file('a', 'One.docx', 'docx')], next: 'p2' },
		});
		provider.signedIn = true;
		const element = await mount(provider);
		button(element, '.more').click();
		await settle();
		expect(names(element)).toEqual(['One.docx', 'Two.docx']);
		expect(element.shadowRoot!.querySelector('.more')).toBeNull();
	});

	it('debounces search and shows locations', async () => {
		vi.useFakeTimers();
		const provider = fakeProvider();
		provider.signedIn = true;
		const element = await mount(provider);
		const field = element.shadowRoot!.querySelector<HTMLInputElement & { value: string }>(
			'office-ui-search',
		)!;
		for (const value of ['r', 're', 'rep']) {
			field.value = value;
			field.dispatchEvent(new Event('input', { bubbles: true }));
			vi.advanceTimersByTime(100);
		}
		expect(provider.searches).toEqual([]);
		vi.advanceTimersByTime(300);
		await settle();
		expect(provider.searches).toEqual(['rep']);
		expect(rows(element)[0]!.textContent).toContain('My files > Old');
	});

	it('disables files of kinds that are not accepted but keeps folders enterable', async () => {
		const provider = fakeProvider();
		provider.signedIn = true;
		const element = await mount(provider, { accept: ['xlsx'] });
		const [dir, docx, txt] = rows(element);
		expect(dir!.getAttribute('aria-disabled')).toBe('false');
		expect(docx!.getAttribute('aria-disabled')).toBe('true');
		expect(txt!.getAttribute('aria-disabled')).toBe('true');
		const listener = vi.fn();
		element.addEventListener('office-cloud-file-open', listener);
		docx!.click();
		await settle();
		expect(listener).not.toHaveBeenCalled();
	});

	it('dispatches office-cloud-file-open with the opened file', async () => {
		const provider = fakeProvider();
		provider.signedIn = true;
		const element = await mount(provider);
		const listener = vi.fn();
		document.body.addEventListener('office-cloud-file-open', listener);
		rows(element)[1]!.click();
		await settle();
		expect(listener).toHaveBeenCalledOnce();
		const detail = (listener.mock.calls[0]![0] as CustomEvent).detail;
		expect(detail.providerId).toBe('fake');
		expect(detail.file.fileName).toBe('Letter.docx');
	});

	it('opens with Enter and moves with the arrow keys', async () => {
		const provider = fakeProvider();
		provider.signedIn = true;
		const element = await mount(provider);
		const listener = vi.fn();
		element.addEventListener('office-cloud-file-open', listener);
		const list = element.shadowRoot!.querySelector('ul')!;
		list.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
		expect(rows(element)[1]!.getAttribute('tabindex')).toBe('0');
		list.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		await settle();
		expect(listener).toHaveBeenCalledOnce();
	});

	it('shows progress, blocks other opens and cancels', async () => {
		let report: CloudOpenOptions['onProgress'];
		let abortSignal: AbortSignal | undefined;
		const provider = fakeProvider({
			open: (entry, options) =>
				new Promise<CloudOpenedFile>((resolve, reject) => {
					report = options?.onProgress;
					abortSignal = options?.signal;
					options?.signal?.addEventListener('abort', () => reject(new Error('aborted')));
					void resolve;
					void entry;
				}),
		});
		provider.signedIn = true;
		const element = await mount(provider);
		const listener = vi.fn();
		element.addEventListener('office-cloud-file-open', listener);
		rows(element)[1]!.click();
		report!(50, 100);
		expect(element.shadowRoot!.querySelector('[role="status"]')!.textContent).toBe(
			'Opening Letter.docx (50%)',
		);
		expect(rows(element)[1]!.getAttribute('aria-busy')).toBe('true');
		button(element, '[data-action="cancel"]').click();
		await settle();
		expect(abortSignal!.aborted).toBe(true);
		expect(element.shadowRoot!.querySelector('[role="status"]')).toBeNull();
		expect(element.shadowRoot!.querySelector('[role="alert"]')!.textContent).toBe('');
		expect(listener).not.toHaveBeenCalled();
	});

	it('shows an inline error and dispatches nothing when opening fails', async () => {
		const provider = fakeProvider({
			open: async () => {
				throw new Error('File is locked');
			},
		});
		provider.signedIn = true;
		const element = await mount(provider);
		const listener = vi.fn();
		element.addEventListener('office-cloud-file-open', listener);
		rows(element)[1]!.click();
		await settle();
		expect(element.shadowRoot!.querySelector('[role="alert"]')!.textContent).toBe('File is locked');
		expect(listener).not.toHaveBeenCalled();
	});

	it('aborts the in-flight listing when the user navigates away', async () => {
		const provider = fakeProvider();
		provider.signedIn = true;
		const element = await mount(provider);
		let pending: AbortSignal | undefined;
		provider.list = (_id, options) => {
			pending ??= options?.signal;
			return new Promise<CloudListing>(() => {});
		};
		rows(element)[0]!.click();
		expect(pending!.aborted).toBe(false);
		button(element, '[data-crumb="0"]').click();
		expect(pending!.aborted).toBe(true);
	});

	it('opens a pasted link when the provider supports it', async () => {
		const openLink = vi.fn(async () => opened(file('l', 'Linked.pptx', 'pptx')));
		const provider = fakeProvider({ openLink });
		provider.signedIn = true;
		const element = await mount(provider);
		const listener = vi.fn();
		element.addEventListener('office-cloud-file-open', listener);
		element.shadowRoot!.querySelector<HTMLInputElement>('.link input')!.value = ' https://x/y ';
		element
			.shadowRoot!.querySelector('form')!
			.dispatchEvent(new Event('submit', { cancelable: true }));
		await settle();
		expect((openLink.mock.calls as unknown as string[][])[0]![0]).toBe('https://x/y');
		expect(listener).toHaveBeenCalledOnce();
	});

	it('reloads when the account changes and localises through labels', async () => {
		const provider = fakeProvider();
		const element = await mount(provider, { labels: { signIn: 'Anmelden' } });
		expect(button(element, '.primary').textContent).toBe('Anmelden');
		provider.signedIn = true;
		provider.notify();
		await settle();
		expect(names(element)).toHaveLength(3);
	});
});
