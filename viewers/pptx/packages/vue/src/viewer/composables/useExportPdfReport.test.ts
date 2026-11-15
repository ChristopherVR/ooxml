// @vitest-environment happy-dom
import type { CompatibilityWarningToast, RunPdfExportOptions } from 'ooxml-ui/pptx';
import type { PptxData, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { computed, effectScope, ref } from 'vue';

const NOTICE: CompatibilityWarningToast = {
	id: 'pdf-export-report',
	code: 'PDF_EXPORT_APPROXIMATED',
	severity: 'info',
	messageKey: 'pptx.pdfExport.summary.message',
};
const RESULT = { diagnostics: [{ code: 'shadowAsPicture' }], slides: [] } as never;

const runs: RunPdfExportOptions[] = [];
const choose = vi.fn();

// The shared export is exercised by its own tests; here only what the binding hands it.
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	runPdfExport: vi.fn(async (options: RunPdfExportOptions) => {
		runs.push(options);
		options.onApproximated?.(RESULT);
		options.notify?.(NOTICE);
		return RESULT;
	}),
	choosePdfExportSelection: ((...args: unknown[]) => choose(...args)) as never,
}));
vi.mock(import('vue-i18n'), async (importOriginal) => ({
	...(await importOriginal()),
	useI18n: (() => ({ t: (key: string) => key })) as never,
}));

const { useExport } = await import('./useExport');
const { useExportWiring } = await import('./useExportWiring');

const slides = () => ref([{ id: 'a' }, { id: 'b', hidden: true }] as PptxSlide[]);

afterEach(() => {
	runs.length = 0;
	choose.mockReset();
});

describe('Vue vector PDF report and notice', () => {
	it('reports to the viewer emit and the per-call onReport, and shows the notice', async () => {
		const onReport = vi.fn();
		const perCall = vi.fn();
		const notify = vi.fn();
		const exporter = useExport({
			slides: slides(),
			canvasSize: ref({ width: 960, height: 540 }),
			rasterizeSlide: vi.fn(),
			mountSlide: vi.fn(),
			onReport,
			notify,
		});
		await exporter.exportPdf({ onReport: perCall, slideIndexes: [1] });
		expect(runs[0]!.vector.slideIndexes).toEqual([1]);
		expect(perCall).toHaveBeenCalledExactlyOnceWith(RESULT);
		expect(onReport).toHaveBeenCalledExactlyOnceWith(RESULT);
		expect(notify).toHaveBeenCalledExactlyOnceWith(NOTICE);
		// The programmatic export never asks for a range.
		expect(choose).not.toHaveBeenCalled();
	});
});

describe('Vue File > Export > PDF options step', () => {
	function wiring() {
		const container = document.createElement('div');
		const onPdfExportReport = vi.fn();
		const notify = vi.fn();
		const scope = effectScope();
		const result = scope.run(() =>
			useExportWiring({
				mergedSlides: ref([]),
				slides: slides(),
				slideCount: computed(() => 2),
				canvasSize: ref({ width: 960, height: 540 }),
				activeSlideIndex: ref(1),
				saveAs: vi.fn(),
				getDeckData: () => ({}) as PptxData,
				onPdfExportReport,
				notify,
				dialogContainer: () => container,
			}),
		)!;
		return { ...result, container, onPdfExportReport, notify, stop: () => scope.stop() };
	}

	it('asks for the range in the viewer root, then exports that selection', async () => {
		choose.mockResolvedValue({ slideIndexes: [1], includeHiddenSlides: true });
		const { onExportPdf, container, onPdfExportReport, notify, stop } = wiring();
		onExportPdf();
		await vi.waitFor(() => expect(runs).toHaveLength(1));
		expect(choose).toHaveBeenCalledWith(
			container,
			expect.objectContaining({
				currentSlide: 1,
				settings: { mode: 'vector', includeHiddenSlides: false },
			}),
		);
		expect(runs[0]!.vector.slideIndexes).toEqual([1]);
		expect(runs[0]!.settings.includeHiddenSlides).toBe(true);
		expect(onPdfExportReport).toHaveBeenCalledWith(RESULT);
		expect(notify).toHaveBeenCalledWith(NOTICE);
		stop();
	});

	it('exports nothing when the options dialog is cancelled', async () => {
		choose.mockResolvedValue(null);
		const { onExportPdf, stop } = wiring();
		onExportPdf();
		await vi.waitFor(() => expect(choose).toHaveBeenCalledOnce());
		await Promise.resolve();
		expect(runs).toHaveLength(0);
		stop();
	});
});
