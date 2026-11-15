/**
 * viewer-export-pdf.test.ts: File > Export > PDF wiring of `ViewerExportService`. The viewer's
 * own UI asks for the range first (the shared options dialog, mounted in the viewer root) and
 * exports the chosen slides; Cancel exports nothing; the export hands the toast stack its notice
 * and the host its `pdfExportReport`.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { afterEach, describe, expect, it, vi } from 'vitest';

const shared = vi.hoisted(() => ({
	runPdfExport: vi.fn(async () => undefined),
	choosePdfExportSelection: vi.fn(),
}));

vi.mock('ooxml-ui/pptx', async (importOriginal) => ({
	...(await importOriginal<Record<string, unknown>>()),
	runPdfExport: shared.runPdfExport,
	choosePdfExportSelection: shared.choosePdfExportSelection,
}));

import { ExportService } from './export.service';
import { LoadContentService } from './load-content.service';
import { PrintService } from './print.service';
import type { ExportHost } from './viewer-export-capture';
import { ViewerExportService } from './viewer-export.service';

function createService(): ViewerExportService {
	const injector = Injector.create({
		providers: [
			{ provide: ExportService, useValue: {} },
			{
				provide: LoadContentService,
				useValue: { canvasSize: () => ({ width: 960, height: 540 }) },
			},
			{ provide: PrintService, useValue: { print: vi.fn() } },
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
		],
	});
	return runInInjectionContext(injector, () => new ViewerExportService());
}

const root = document.createElement('div');
const notify = vi.fn();
const onPdfExportReport = vi.fn();

function host(): ExportHost {
	return {
		activeSlideIndex: signal(1),
		slideCount: () => 3,
		mergedSlides: () => [],
		slides: () => [{}, { hidden: true }, {}],
		resolveStage: () => document.createElement('div'),
		pdfExportSettings: () => ({ mode: 'vector', includeHiddenSlides: false }),
		pdfOptionsContainer: () => root,
		notify,
		onPdfExportReport,
	};
}

afterEach(() => vi.clearAllMocks());

describe('viewerExportService PDF export', () => {
	it('asks for the range in the viewer root and exports the chosen slides', async () => {
		shared.choosePdfExportSelection.mockResolvedValue({
			slideIndexes: [1],
			includeHiddenSlides: true,
		});
		const svc = createService();
		svc.bind(host());
		await svc.exportPdfWithOptions();
		expect(shared.choosePdfExportSelection).toHaveBeenCalledWith(
			root,
			expect.objectContaining({
				currentSlide: 1,
				settings: { mode: 'vector', includeHiddenSlides: false },
				slides: [{}, { hidden: true }, {}],
			}),
		);
		expect(shared.runPdfExport).toHaveBeenCalledOnce();
		const options = shared.runPdfExport.mock.calls[0]![0] as unknown as {
			vector: Record<string, unknown>;
			notify: unknown;
			onApproximated: unknown;
		};
		expect(options.vector).toMatchObject({ slideIndexes: [1], includeHiddenSlides: true });
		// The notice goes to the toast stack, the report to the `pdfExportReport` output.
		expect(options.notify).toBe(notify);
		expect(options.onApproximated).toBe(onPdfExportReport);
	});

	it('exports nothing when the options dialog is cancelled', async () => {
		shared.choosePdfExportSelection.mockResolvedValue(null);
		const svc = createService();
		svc.bind(host());
		await svc.exportPdfWithOptions();
		expect(shared.runPdfExport).not.toHaveBeenCalled();
	});

	it('exports every slide without a dialog when called directly', async () => {
		const svc = createService();
		svc.bind(host());
		await svc.exportPdf();
		expect(shared.choosePdfExportSelection).not.toHaveBeenCalled();
		const options = shared.runPdfExport.mock.calls[0]![0] as unknown as {
			vector: Record<string, unknown>;
		};
		expect(options.vector.slideIndexes).toBeUndefined();
	});
});
