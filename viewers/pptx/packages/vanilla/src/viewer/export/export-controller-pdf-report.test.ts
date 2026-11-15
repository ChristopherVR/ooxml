import type { RunPdfExportOptions, VectorPdfExportResult } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import { createExportController } from './export-controller';

/**
 * The vanilla PDF export hands the shared `runPdfExport` everything it reports through: the
 * per-call `onReport`, the viewer-level `onPdfExportReport`, the toast `notify` and the slide
 * selection the options step chose (including the current slide's explicit index).
 */
const { runPdfExport } = vi.hoisted(() => ({
	runPdfExport: vi.fn((_options: RunPdfExportOptions) => Promise.resolve(undefined)),
}));
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	runPdfExport,
}));

function store() {
	const state = createStore(createInitialViewerState());
	state.set({
		slides: [0, 1, 2].map((i) => ({
			id: `s${i}`,
			rId: `r${i}`,
			slideNumber: i + 1,
			elements: [],
		})) as unknown as PptxSlide[],
		canvasSize: { width: 960, height: 540 },
	});
	return state;
}

describe('vanilla vector PDF report wiring', () => {
	afterEach(() => vi.clearAllMocks());

	it('passes the selection, notify and both report callbacks to runPdfExport', async () => {
		const notify = vi.fn();
		const onPdfExportReport = vi.fn();
		const onReport = vi.fn();
		const controller = createExportController({
			store: store(),
			rasterizeSlide: vi.fn(),
			mountSlide: vi.fn(),
			notify,
			onPdfExportReport,
		});
		await controller.exportPdf({ slideIndexes: [2], includeHiddenSlides: true, onReport });
		const [options] = runPdfExport.mock.calls[0]!;
		expect(options.vector.slideIndexes).toEqual([2]);
		expect(options.settings).toEqual({ mode: 'vector', includeHiddenSlides: true });
		expect(options.notify).toBe(notify);
		const result = { diagnostics: [{ code: 'shadowAsPicture' }] } as VectorPdfExportResult;
		options.onApproximated!(result);
		expect(onReport).toHaveBeenCalledWith(result);
		expect(onPdfExportReport).toHaveBeenCalledWith(result);
	});

	it('passes a from-to range through unchanged', async () => {
		const controller = createExportController({
			store: store(),
			rasterizeSlide: vi.fn(),
			mountSlide: vi.fn(),
		});
		await controller.exportPdf({ from: 2, to: 3 });
		const [options] = runPdfExport.mock.calls[0]!;
		expect(options.vector).toMatchObject({ from: 2, to: 3 });
		expect(options.vector.slideIndexes).toBeUndefined();
	});
});
