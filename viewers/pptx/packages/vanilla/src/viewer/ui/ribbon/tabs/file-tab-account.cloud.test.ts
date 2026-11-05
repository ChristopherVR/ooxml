/**
 * File > Account with `fileProviders`: the shared connected services list gets the providers and
 * translated labels, and nothing extra renders without providers.
 */
import type { CloudFileProvider } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../../i18n';
import { renderAccountPage } from './file-tab-account';

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	getLocalStorageUsageSummary: async () => ({ presentationCount: 0, totalBytes: 0 }),
}));

const t = createTranslator('en');

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

afterEach(() => {
	document.body.replaceChildren();
});

interface ServicesElement extends HTMLElement {
	providers: unknown[];
	labels: Record<string, string>;
}

describe('File > Account connected services', () => {
	it('renders no connected services without providers', () => {
		const main = document.createElement('div');
		renderAccountPage(document, t, main, undefined);
		expect(main.querySelector('pptx-ui-connected-services')).toBeNull();
	});

	it('hands the fileProviders and translated labels to the shared list', () => {
		const main = document.createElement('div');
		renderAccountPage(document, t, main, undefined, undefined, [provider]);
		const list = main.querySelector<ServicesElement>('pptx-ui-connected-services')!;
		expect(list.providers).toEqual([provider]);
		expect(list.labels['heading']).toBe('Connected services');
		expect(list.labels['connect']).toBe('Connect');
	});
});
