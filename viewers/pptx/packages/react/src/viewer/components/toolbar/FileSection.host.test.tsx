// @vitest-environment happy-dom
/**
 * The File backstage covers its viewer host, not the page around it, and Escape pressed inside
 * it closes it while keys pressed elsewhere on the page are left alone.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => key, locale: 'en' }),
}));

const { FileSection } = await import('./FileSection');
type FileSectionProps = import('./file-backstage-parts').FileSectionProps;

let container: HTMLDivElement;
let outside: HTMLButtonElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	outside = document.createElement('button');
	container = document.createElement('div');
	container.setAttribute('data-pptx-theme-root', '');
	document.body.append(outside, container);
	vi.spyOn(container, 'getBoundingClientRect').mockReturnValue({
		left: 160,
		top: 40,
		width: 1000,
		height: 700,
		right: 1160,
		bottom: 740,
		x: 160,
		y: 40,
	} as DOMRect);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	outside.remove();
	vi.restoreAllMocks();
});

function props(onClose: () => void): FileSectionProps {
	const noop = () => {};
	return {
		onClose,
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
	} as FileSectionProps;
}

const escape = (target: Element): void => {
	target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
};

describe('fileSection host containment', () => {
	it('sits on the viewer host box and closes on Escape pressed inside it', () => {
		const onClose = vi.fn();
		act(() => root.render(<FileSection {...props(onClose)} />));
		const panel = container.querySelector<HTMLElement>('[role="dialog"]')!;
		expect([
			panel.style.left,
			panel.style.top,
			panel.style.width,
			panel.style.height,
		]).toStrictEqual(['160px', '40px', '1000px', '700px']);
		escape(outside);
		expect(onClose).not.toHaveBeenCalled();
		escape(panel.querySelector('button')!);
		expect(onClose).toHaveBeenCalledOnce();
	});
});
