/**
 * Save-side glue between an edited chart and the workbook it embeds: checks that the chart's
 * `c:externalData` resolves to an embedded `.xlsx` part of the package being saved, hands the
 * collected writes to {@link writeChartWorkbookUpdates} and stores the result back. Every way the
 * reference can fail (no linked workbook, an external or non-xlsx target, a missing or unreadable
 * part, writes the range parser could not place) degrades safely: the chart cache edit stands and
 * an issue is reported instead of throwing. An atomic save throws instead, so a caller cannot
 * publish mismatched chart caches and workbook cells.
 *
 * Shared by PowerPoint (`pptx/core/core/runtime/chart-external-workbook-save.ts`, which adds the
 * slide context to each issue) and Word (`docx/chart-save.ts`).
 *
 * @module chart/workbook/package-save
 */
import { writeChartWorkbookUpdates, type ChartWorkbookWrite } from './writer';

/** The package being saved, as far as the workbook write-back needs it. */
export interface ChartWorkbookPackage {
	read(partName: string): Promise<Uint8Array | undefined>;
	write(partName: string, bytes: Uint8Array): void;
}

/** Where the chart's data lives: the resolved `c:externalData` target. */
export interface ChartWorkbookTarget {
	/** The package part name of the workbook, resolved against the chart part. */
	partName?: string;
	/** `TargetMode="External"`: the workbook is outside the package. */
	external?: boolean;
}

export type ChartWorkbookIssueCode =
	| 'CHART_EXTERNAL_DATA_WRITEBACK_UNSUPPORTED'
	| 'CHART_EMBEDDED_WORKBOOK_MISSING'
	| 'CHART_EMBEDDED_WORKBOOK_UNREADABLE'
	| 'CHART_EMBEDDED_WORKBOOK_PARTIAL_WRITEBACK';

/** Why some or all of a chart's data stayed cache-only. */
export interface ChartWorkbookIssue {
	code: ChartWorkbookIssueCode;
	message: string;
	severity: 'info' | 'warning';
	/** The chart part, or the workbook part once it resolved. */
	partName: string;
}

export interface ChartWorkbookSaveOptions {
	atomic?: boolean;
	report?(issue: ChartWorkbookIssue): void;
}

/** Whether `partName` is an embedded `.xlsx` package this module can update. */
export const isEmbeddedXlsxPart = (partName: string): boolean =>
	partName.includes('embeddings/') && /\.xlsx$/iu.test(partName);

/**
 * Rewrites the chart's embedded workbook to match its edited names, categories and values.
 * Returns true when every write landed (or there was nothing to write).
 */
export async function saveChartWorkbookUpdates(
	pkg: ChartWorkbookPackage,
	chartPartName: string,
	target: ChartWorkbookTarget | undefined,
	writes: readonly ChartWorkbookWrite[],
	options: ChartWorkbookSaveOptions = {},
): Promise<boolean> {
	const { atomic = false } = options;
	const report = (issue: ChartWorkbookIssue) => options.report?.(issue);
	if (atomic && target && (!target.partName || target.external))
		throw new Error('Chart label replacement cannot update an external or missing workbook.');
	// No data edits, or a chart without linked data (most charts): nothing to reconcile.
	if (writes.length === 0 || !target?.partName) return writes.length === 0;
	const partName = target.partName;
	if (!isEmbeddedXlsxPart(partName) || target.external) {
		if (atomic) throw new Error('Chart label replacement requires an embedded .xlsx workbook.');
		report({
			code: 'CHART_EXTERNAL_DATA_WRITEBACK_UNSUPPORTED',
			message:
				'Chart data changes were saved to the cached values only: the linked data source is external or not an embedded .xlsx workbook, so it was not updated.',
			severity: 'info',
			partName: chartPartName,
		});
		return false;
	}
	let bytes: Uint8Array | undefined;
	try {
		bytes = await pkg.read(partName);
	} catch {
		if (atomic) throw new Error('Chart label replacement workbook cannot be read.');
		report({
			code: 'CHART_EMBEDDED_WORKBOOK_UNREADABLE',
			message:
				'Chart data changes were saved to the cached values only: the referenced embedded workbook could not be read.',
			severity: 'warning',
			partName,
		});
		return false;
	}
	if (!bytes) {
		if (atomic) throw new Error('Chart label replacement workbook is missing.');
		report({
			code: 'CHART_EMBEDDED_WORKBOOK_MISSING',
			message:
				'Chart data changes were saved to the cached values only: the referenced embedded workbook was not found in the package.',
			severity: 'warning',
			partName,
		});
		return false;
	}
	const result = await writeChartWorkbookUpdates(bytes, writes, { atomic });
	if (result.bytes) pkg.write(partName, result.bytes);
	if (result.unresolved > 0)
		report({
			code: 'CHART_EMBEDDED_WORKBOOK_PARTIAL_WRITEBACK',
			message: `${result.unresolved} chart data update(s) could not be matched to a cell in the linked workbook and were saved to the cached values only.`,
			severity: 'info',
			partName,
		});
	return result.unresolved === 0;
}
