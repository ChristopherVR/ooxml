/**
 * The public `exportSlidesVectorPdf` (see `pdf-vector-export.ts`). The vector PDF writer loads on
 * first use, so importing the viewer or its export helpers does not download it.
 */
import type { VectorPdfExportOptions, VectorPdfExportResult } from './pdf-vector-export';

export type {
	MountedPdfSlide,
	VectorPdfExportOptions,
	VectorPdfExportResult,
	VectorPdfSlideReport,
} from './pdf-vector-export';

export async function exportSlidesVectorPdf(
	options: VectorPdfExportOptions,
): Promise<VectorPdfExportResult> {
	const { exportSlidesVectorPdf: run } = await import('./pdf-vector-export');
	return run(options);
}
