import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { ExportControllerDeps } from './export';
import { createTranslator } from './i18n';
import { createElementRendererRegistry } from './render';
import { createInitialViewerState, createStore } from './state';

/**
 * File > Export > PDF from the vanilla UI asks for the range first (the shared
 * `choosePdfExportSelection`), then exports what was chosen; Cancel exports nothing, and the
 * programmatic `exportPdf()` never asks. The viewer's report callback and toast `notify` reach
 * the export controller.
 */
const { runPdf, choosePdfExportSelection, createExportController } = vi.hoisted(() => ({
	runPdf: vi.fn(),
	choosePdfExportSelection: vi.fn(),
	createExportController: vi.fn((_deps: ExportControllerDeps) => ({})),
}));
vi.mock(
	import('./export'),
	() =>
		({
			createRasterizeSlide: vi.fn(() => ({ mountSlide: vi.fn(), destroy: vi.fn() })),
			createExportController,
			createExportProgressModal: vi.fn(() => ({ open: vi.fn(), update: vi.fn(), close: vi.fn() })),
			createExportProgressUi: vi.fn(() => ({ runPdf, cancel: vi.fn() })),
		}) as never,
);
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	choosePdfExportSelection,
}));

const { createExportLifecycle } = await import('./export-lifecycle');

function deps(extra: Record<string, unknown> = {}) {
	const store = createStore(createInitialViewerState());
	store.set({
		slides: [{ id: 'a' }, { id: 'b', hidden: true }] as unknown as PptxSlide[],
		currentSlide: 1,
	});
	const host = document.createElement('div');
	return {
		host,
		deps: {
			doc: document,
			container: document.createElement('div'),
			getModalHost: () => host,
			store,
			registry: createElementRendererRegistry(),
			getTranslator: () => createTranslator(),
			getSmartArt3D: () => false,
			getSurfaceChart3D: () => false,
			getBarChart3D: () => false,
			getLineChart3D: () => false,
			getAreaChart3D: () => false,
			getPieChart3D: () => false,
			getImageResolutionScale: () => 1,
			getPdfExportSettings: () => ({ mode: 'vector' as const, includeHiddenSlides: false }),
			...extra,
		},
	};
}

describe('vanilla Export > PDF options step', () => {
	afterEach(() => vi.clearAllMocks());

	it('asks in the viewer root, then exports the chosen slides', async () => {
		const { host, deps: d } = deps();
		choosePdfExportSelection.mockResolvedValue({ slideIndexes: [1], includeHiddenSlides: true });
		await createExportLifecycle(d).exportPdfFromUi();
		const [container, input] = choosePdfExportSelection.mock.calls[0]!;
		expect(container).toBe(host);
		expect(input).toMatchObject({
			settings: { mode: 'vector', includeHiddenSlides: false },
			currentSlide: 1,
		});
		expect(input.slides).toHaveLength(2);
		expect(runPdf).toHaveBeenCalledExactlyOnceWith({
			includeHiddenSlides: true,
			slideIndexes: [1],
		});
	});

	it('exports nothing when the options step is cancelled', async () => {
		choosePdfExportSelection.mockResolvedValue(null);
		await createExportLifecycle(deps().deps).exportPdfFromUi();
		expect(runPdf).not.toHaveBeenCalled();
	});

	it('never asks for the programmatic exportPdf()', async () => {
		await createExportLifecycle(deps().deps).exportPdf({ from: 1, to: 1 });
		expect(choosePdfExportSelection).not.toHaveBeenCalled();
		expect(runPdf).toHaveBeenCalledWith({ from: 1, to: 1 });
	});

	it('hands the report callback and the toast notify to the controller', () => {
		const onPdfExportReport = vi.fn();
		const notify = vi.fn();
		createExportLifecycle(deps({ onPdfExportReport, notify }).deps);
		const [controllerDeps] = createExportController.mock.calls[0]!;
		const result = { diagnostics: [] } as never;
		controllerDeps.onPdfExportReport!(result);
		expect(onPdfExportReport).toHaveBeenCalledWith(result);
		const toast = { id: 'pdf-export-report' } as never;
		controllerDeps.notify!(toast);
		expect(notify).toHaveBeenCalledWith(toast);
	});
});
