// @vitest-environment happy-dom
/**
 * File > Export > PDF in React: the shared options step (`choosePdfExportSelection`) runs first,
 * in the viewer root, and its range, the toast pusher and the host's `onPdfExportReport` reach
 * the export. Cancel exports nothing.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ExportHandlersResult, UseExportHandlersInput } from './export-handler-types';
import { useExportHandlers } from './useExportHandlers';

const pdfMock = vi.hoisted(() => ({
	exportPresentationPdf: vi.fn<typeof import('../utils/export-pdf-vector').exportPresentationPdf>(
		async () => undefined,
	),
}));
vi.mock(import('../utils/export-pdf-vector'), () => pdfMock);
vi.mock(import('../utils/dom-helpers'), () => ({ downloadBlob: vi.fn() }));

let root: Root | undefined;
let host: HTMLDivElement | undefined;
let api: ExportHandlersResult;

function Harness({ input }: { input: UseExportHandlersInput }): null {
	api = useExportHandlers(input);
	return null;
}

const slide = (id: string, hidden = false) =>
	({ id, elements: [], backgroundColor: '#fff', notes: '', hidden }) as unknown as PptxSlide;

function setup(extra: Partial<UseExportHandlersInput> = {}) {
	const container = document.createElement('div');
	document.body.append(container);
	const slides = [slide('a'), slide('b', true), slide('c')];
	const input: UseExportHandlersInput = {
		slides,
		activeSlide: slides[1],
		activeSlideIndex: 1,
		templateElementsBySlideId: {},
		filePath: 'deck.pptx',
		canvasStageRef: { current: document.createElement('div') },
		setActiveSlideIndex: vi.fn(),
		serializeSlides: vi.fn(async () => null),
		headerFooter: {},
		presentationProperties: {},
		customShows: [],
		sections: [],
		coreProperties: undefined,
		appProperties: undefined,
		customProperties: [],
		tagCollections: [],
		notesMaster: undefined,
		handoutMaster: undefined,
		theme: undefined,
		canvasSize: { width: 960, height: 540 },
		pdfExportSettings: { mode: 'vector', includeHiddenSlides: false },
		pdfExportContainerRef: { current: container },
		notifyPdfExport: vi.fn(),
		onPdfExportReport: vi.fn(),
		...extra,
	};
	host = document.createElement('div');
	document.body.append(host);
	root = createRoot(host);
	act(() => root!.render(createElement(Harness, { input })));
	return { input, container };
}

type Shadowed = HTMLElement & { shadowRoot: ShadowRoot };
const footerButton = (dialog: Element, id: string) =>
	(
		dialog.querySelector('office-ui-dialog-footer') as Shadowed
	).shadowRoot.querySelector<HTMLButtonElement>(`button[data-action="${id}"]`)!;

beforeEach(async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	// Registers the shared web controls (the PDF options dialog among them).
	await import('ooxml-ui/pptx');
});

afterEach(() => {
	act(() => root?.unmount());
	host?.remove();
	document.body.replaceChildren();
	root = undefined;
	host = undefined;
	vi.clearAllMocks();
	vi.unstubAllGlobals();
});

describe('pdf export options step (react)', () => {
	it('asks for the range in the viewer root and exports the chosen slides', async () => {
		const { input, container } = setup();
		let pending!: Promise<void>;
		act(() => {
			pending = api.handleExportPdf();
		});
		const dialog = container.querySelector('pptx-ui-pdf-export-dialog')!;
		expect(dialog).not.toBeNull();
		expect(pdfMock.exportPresentationPdf).not.toHaveBeenCalled();
		dialog.querySelector<HTMLElement>('[data-testid="pptx-pdf-export-range-current"]')!.click();
		await act(async () => {
			footerButton(dialog, 'insert').click();
			await pending;
		});
		expect(container.querySelector('pptx-ui-pdf-export-dialog')).toBeNull();
		const options = pdfMock.exportPresentationPdf.mock.calls[0]![4];
		expect(options.selection).toStrictEqual({ slideIndexes: [1], includeHiddenSlides: true });
		expect(options.notify).toBe(input.notifyPdfExport);
		expect(options.onApproximated).toBe(input.onPdfExportReport);
	});

	it('exports nothing when the options step is cancelled', async () => {
		const { container } = setup();
		let pending!: Promise<void>;
		act(() => {
			pending = api.handleExportPdf();
		});
		const dialog = container.querySelector('pptx-ui-pdf-export-dialog')!;
		await act(async () => {
			footerButton(dialog, 'cancel').click();
			await pending;
		});
		expect(pdfMock.exportPresentationPdf).not.toHaveBeenCalled();
		expect(api.exportModalOpen).toBeFalsy();
	});

	it('asks nothing for the picture export', async () => {
		const { container } = setup({
			pdfExportSettings: { mode: 'pictures', includeHiddenSlides: false },
		});
		await act(async () => api.handleExportPdf());
		expect(container.childElementCount).toBe(0);
		expect(pdfMock.exportPresentationPdf.mock.calls[0]![4].selection).toStrictEqual({});
	});
});
