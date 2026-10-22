// @vitest-environment happy-dom
/**
 * Host-owned storage in the File backstage: the nav's Save item is File > Save
 * (`onSave`, reported to the host's storage as `save`), not the Save As .pptx
 * card; and the Recent list reads the host's recovery store from the viewer's
 * storage context.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const listBackstageRecentFiles = vi.fn(async () => []);

vi.mock(import('ooxml-ui/pptx'), async (original) => ({
	...(await original()),
	listBackstageRecentFiles,
}));

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const { FileSection } = await import('./FileSection');
const { ViewerStorageContext } = await import('../viewer-storage-context');
type FileSectionProps = import('./file-backstage-parts').FileSectionProps;

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	listBackstageRecentFiles.mockClear();
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function props(overrides: Partial<FileSectionProps>): FileSectionProps {
	const noop = () => {};
	return {
		onClose: noop,
		onCreatePresentation: noop,
		onSaveAsPptx: vi.fn(),
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

describe('FileSection with host-owned storage', () => {
	it('runs onSave from the Save nav item, not the Save As .pptx handler', () => {
		const onSave = vi.fn();
		const p = props({ onSave });
		act(() => root.render(<FileSection {...p} />));
		const save = [
			...container.querySelectorAll<HTMLButtonElement>('[data-pptx-backstage-nav-item]'),
		].find((button) => button.textContent?.includes('ave') && !button.textContent.includes('As'));
		act(() => save!.click());
		expect(onSave).toHaveBeenCalledOnce();
		expect(p.onSaveAsPptx).not.toHaveBeenCalled();
	});

	it('lists Recent files through the viewer storage', () => {
		const storage = { save: async () => undefined };
		act(() =>
			root.render(
				<ViewerStorageContext.Provider value={storage}>
					<FileSection {...props({})} />
				</ViewerStorageContext.Provider>,
			),
		);
		expect(listBackstageRecentFiles).toHaveBeenCalledWith(expect.any(Function), undefined, storage);
	});
});
