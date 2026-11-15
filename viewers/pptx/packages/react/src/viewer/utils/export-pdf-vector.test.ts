// @vitest-environment happy-dom
/**
 * `exportPresentationPdf` hands the options step's range to the vector export and the toast
 * pusher and host report to the shared `runPdfExport`, which shows the notice.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';

const shared = vi.hoisted(() => ({ runPdfExport: vi.fn(async () => undefined) }));
vi.mock(import('ooxml-ui/pptx'), async (original) => ({
	...(await original()),
	runPdfExport: shared.runPdfExport as never,
}));
vi.mock(import('../../lib/canvas-export'), () => ({ renderToCanvas: vi.fn() }));
vi.mock(import('./export-pdf'), () => ({ exportAllSlidesAsPdf: vi.fn() }));

const { exportPresentationPdf } = await import('./export-pdf-vector');

afterEach(() => vi.clearAllMocks());

describe('exportPresentationPdf', () => {
	it('forwards the selection, the notice and the report to runPdfExport', async () => {
		const notify = vi.fn();
		const onApproximated = vi.fn();
		const setActive = vi.fn();
		await exportPresentationPdf({ current: document.createElement('div') }, setActive, 2, 'a.pdf', {
			settings: { mode: 'vector', includeHiddenSlides: false },
			slides: [{}, {}, {}],
			canvasSize: { width: 960, height: 540 },
			selection: { from: 2, to: 3, includeHiddenSlides: true },
			notify,
			onApproximated,
		});
		const options = (shared.runPdfExport.mock.calls[0] as unknown as [Record<string, never>])[0];
		expect(options.vector).toMatchObject({ from: 2, to: 3, includeHiddenSlides: true });
		expect(options.notify).toBe(notify);
		expect(options.onApproximated).toBe(onApproximated);
		// The live slide is put back after the export.
		expect(setActive).toHaveBeenLastCalledWith(2);
	});
});
