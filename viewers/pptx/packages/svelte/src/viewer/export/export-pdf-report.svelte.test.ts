import type {
	CompatibilityWarningToast,
	RunPdfExportOptions,
	VectorPdfExportResult,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { runQuickAccessCommand } from '../state/quick-access-commands';
import type { QuickAccessCommandDeps } from '../state/quick-access-commands';
import { ExportController } from './export-controller.svelte';
import type { ExportControllerDeps } from './export-controller.svelte';
import { ExportUiState } from './export-ui.svelte';

/**
 * File > Export > PDF in Svelte: the options step's selection reaches the vector export, the
 * approximation notice goes to the toast stack, and the host's `onpdfexportreport` hears every
 * report alongside a call's own `onReport`.
 */
const { runPdfExportMock } = vi.hoisted(() => ({ runPdfExportMock: vi.fn() }));
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	runPdfExport: runPdfExportMock,
}));

const RESULT = { diagnostics: [{ code: 'shadowAsPicture' }] } as unknown as VectorPdfExportResult;
const TOAST = { id: 'pdf-export-report' } as CompatibilityWarningToast;

function controller(overrides: Partial<ExportControllerDeps> = {}): ExportController {
	return new ExportController({
		getSlideCount: () => 3,
		getCurrent: () => 1,
		getCanvasSize: () => ({ width: 960, height: 540 }),
		getSlides: () => [{}, {}, {}] as unknown as PptxSlide[],
		rasterizeSlide: vi.fn(),
		mountSlide: vi.fn(),
		...overrides,
	});
}

/** The options the controller handed `runPdfExport`, after it ran the report hooks. */
async function lastRun(): Promise<RunPdfExportOptions> {
	const options = runPdfExportMock.mock.calls.at(-1)![0] as RunPdfExportOptions;
	options.onApproximated?.(RESULT);
	options.notify?.(TOAST);
	return options;
}

describe('svelte PDF export report', () => {
	afterEach(() => vi.clearAllMocks());

	it('tells the call, the host and the toast stack what the vector PDF approximated', async () => {
		const notify = vi.fn();
		const onPdfExportReport = vi.fn();
		const onReport = vi.fn();
		await controller({ notify, onPdfExportReport }).exportPdf({ onReport });
		await lastRun();
		expect(onReport).toHaveBeenCalledExactlyOnceWith(RESULT);
		expect(onPdfExportReport).toHaveBeenCalledExactlyOnceWith(RESULT);
		expect(notify).toHaveBeenCalledExactlyOnceWith(TOAST);
	});

	it('passes the chosen slides to the vector export', async () => {
		await controller().exportPdf({ slideIndexes: [1], includeHiddenSlides: true });
		const options = await lastRun();
		expect(options.settings.includeHiddenSlides).toBe(true);
		expect(options.vector.slideIndexes).toEqual([1]);
	});

	it('asks for the options before the UI export and stops on Cancel', async () => {
		const exporter = controller();
		const choose = vi.fn().mockResolvedValueOnce(null).mockResolvedValueOnce({ from: 2, to: 3 });
		const ui = new ExportUiState({
			controller: exporter,
			getTranslator: () => (key: string) => key,
			choosePdfSelection: choose,
		});
		await ui.runPdf();
		expect(runPdfExportMock).not.toHaveBeenCalled();
		await ui.runPdf();
		const options = await lastRun();
		expect(choose).toHaveBeenCalledTimes(2);
		expect(options.vector).toMatchObject({ from: 2, to: 3 });
	});

	it('runs Quick Access > Export PDF through the options step, not the bare API', () => {
		const runPdf = vi.fn(() => Promise.resolve());
		const exportPdf = vi.fn(() => Promise.resolve());
		runQuickAccessCommand('exportPdf', {
			exportingApi: { exportPdf },
			runPdf,
		} as unknown as QuickAccessCommandDeps);
		expect(runPdf).toHaveBeenCalledOnce();
		expect(exportPdf).not.toHaveBeenCalled();
	});
});
