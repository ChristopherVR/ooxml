// @vitest-environment happy-dom
/**
 * File > Open with cloud places: the Open page lists "This device" and one place per provider,
 * shows the provider's file browser, and hands a file opened there to the viewer.
 */
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => key, locale: 'en' }),
}));

const { FileSection } = await import('./FileSection');
const { ViewerCloudFilesContext } = await import('../viewer-cloud-files-context');
type FileSectionProps = import('./file-backstage-parts').FileSectionProps;

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

function props(overrides: Partial<FileSectionProps>): FileSectionProps {
	const noop = () => {};
	return {
		onClose: noop,
		onCreatePresentation: noop,
		onSaveAsPptx: noop,
		onExportPng: noop,
		onExportPdf: noop,
		onExportVideo: noop,
		onExportGif: noop,
		onExportJson: noop,
		onSaveAsPpsx: noop,
		onSaveAsPptm: noop,
		onSaveAsPpt: noop,
		hasMacros: false,
		onCopySlideAsImage: noop,
		onPrint: noop,
		...overrides,
	} as FileSectionProps;
}

function openPage(): void {
	const open = [
		...container.querySelectorAll<HTMLButtonElement>('[data-pptx-backstage-nav-item]'),
	].find((button) => button.textContent === 'pptx.backstage.nav.open');
	act(() => open!.click());
}

describe('FileSection open places', () => {
	it('keeps the Open page as it was without providers', () => {
		act(() => root.render(<FileSection {...props({})} />));
		openPage();
		expect(container.querySelector('[data-pptx-open-place]')).toBeNull();
		expect(container.textContent).toContain('pptx.backstage.browseDevice');
	});

	it('lists this device and the provider, and opens a provider file', () => {
		const open = vi.fn();
		const onClose = vi.fn();
		act(() =>
			root.render(
				<ViewerCloudFilesContext.Provider value={{ providers: [provider], open, clear: vi.fn() }}>
					<FileSection {...props({ onClose })} />
				</ViewerCloudFilesContext.Provider>,
			),
		);
		openPage();
		const places = [...container.querySelectorAll<HTMLButtonElement>('[data-pptx-open-place]')];
		expect(places.map((place) => place.textContent)).toEqual([
			'pptx.backstage.thisDevice',
			'OneDrive',
		]);
		expect(container.querySelector('pptx-ui-cloud-files')).toBeNull();
		act(() => places[1]!.click());
		expect(container.textContent).not.toContain('pptx.backstage.browseDevice');
		const file = { fileName: 'Deck.pptx' } as CloudOpenedFile;
		act(() => {
			container
				.querySelector('pptx-ui-cloud-files')!
				.dispatchEvent(
					new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
				);
		});
		expect(open).toHaveBeenCalledWith(file);
		expect(onClose).toHaveBeenCalled();
	});
});
