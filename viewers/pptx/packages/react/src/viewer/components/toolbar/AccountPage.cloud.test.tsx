// @vitest-environment happy-dom
/**
 * File > Account with `fileProviders`: the connected services list appears below the other
 * sections, and nothing extra renders without providers.
 */
import type { CloudFileProvider } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => `t:${key}`, locale: 'en' }),
}));

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	getLocalStorageUsageSummary: async () => ({ presentationCount: 0, totalBytes: 0 }),
}));

const { AccountPage } = await import('./AccountPage');
const { ViewerCloudFilesContext } = await import('../viewer-cloud-files-context');

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

type ServicesElement = HTMLElement & {
	providers: readonly CloudFileProvider[];
	labels: Record<string, string>;
};

describe('AccountPage connected services', () => {
	it('renders no connected services without providers', () => {
		act(() => root.render(<AccountPage />));
		expect(container.querySelector('pptx-ui-connected-services')).toBeNull();
	});

	it('hands the fileProviders and translated labels to the shared list', () => {
		act(() =>
			root.render(
				<ViewerCloudFilesContext.Provider
					value={{ providers: [provider], open: vi.fn(), clear: vi.fn() }}
				>
					<AccountPage />
				</ViewerCloudFilesContext.Provider>,
			),
		);
		const list = container.querySelector<ServicesElement>('pptx-ui-connected-services')!;
		expect(list.providers).toEqual([provider]);
		expect(list.labels.heading).toBe('t:pptx.cloudFiles.connectedHeading');
		expect(list.labels.connect).toBe('t:pptx.cloudFiles.connectedConnect');
	});
});
