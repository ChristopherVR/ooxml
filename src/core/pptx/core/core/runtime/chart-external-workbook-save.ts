/**
 * Save-side glue between an edited chart's cache-write loop and the shared embedded-workbook
 * writer. The resolution, write-back and degrade-to-warning policy live in the format-neutral
 * chart area (`chart/workbook/package-save.ts`, shared with Word); this module resolves the
 * pptx relationship target inside the package zip being saved and adds the slide context to
 * each warning. Opt-in atomic label edits throw instead, so callers cannot publish mismatched
 * chart caches and workbook cells.
 *
 * @module chart-external-workbook-save
 */

import type JSZip from 'jszip';

import { saveChartWorkbookUpdates } from '../../../../chart/workbook/package-save';
import {
	restoreChartWorkbookCells,
	type ChartWorkbookCellChange,
} from '../../../../chart/workbook/writer';
import type { CompatibilityWarningInput } from '../../services/PptxCompatibilityService';
import type { PptxExternalData } from '../../types';
import type { PptxChartWorkbookWrite } from '../../utils/chart-xlsx-writer';

/** The runtime primitives this module needs, kept explicit for testability. */
export interface ChartExternalWorkbookSaveDeps {
	zip: JSZip;
	resolveImagePath(basePath: string, target: string): string;
	reportWarning(warning: CompatibilityWarningInput): void;
}

/** The cells one write-back changed in an embedded workbook. */
export interface ChartWorkbookCellsWritten {
	partName: string;
	changes: readonly ChartWorkbookCellChange[];
}

/** Whether every write landed, and the cells the writes changed, when any did. */
export interface ChartExternalWorkbookSaveResult {
	complete: boolean;
	written?: ChartWorkbookCellsWritten;
}

/**
 * Puts back the cells `written` changed wherever they still hold what it wrote, so a write-back
 * of another chart sharing the workbook, made since, stays.
 */
export async function restoreChartWorkbookWrite(
	zip: JSZip,
	written: ChartWorkbookCellsWritten,
): Promise<void> {
	const bytes = await zip.file(written.partName)?.async('uint8array');
	const restored = bytes && (await restoreChartWorkbookCells(bytes, written.changes));
	if (restored) zip.file(written.partName, restored);
}

/**
 * Rewrite a chart's embedded workbook to match its edited series/category/
 * name data. Atomic label edits throw on unsupported sources; other edits report warnings.
 */
export async function saveChartExternalWorkbookUpdates(
	deps: ChartExternalWorkbookSaveDeps,
	chartPartPath: string,
	slidePath: string,
	externalData: PptxExternalData | undefined,
	writes: readonly PptxChartWorkbookWrite[],
	atomic = false,
): Promise<ChartExternalWorkbookSaveResult> {
	const target = externalData
		? {
				...(externalData.targetPath
					? { partName: deps.resolveImagePath(chartPartPath, externalData.targetPath) }
					: {}),
				...(externalData.isExternal ? { external: true } : {}),
			}
		: undefined;
	let written: ChartWorkbookCellsWritten | undefined;
	const complete = await saveChartWorkbookUpdates(
		{
			read: async (partName) => await deps.zip.file(partName)?.async('uint8array'),
			write: (partName, bytes, changes) => {
				deps.zip.file(partName, bytes);
				written = { partName, changes };
			},
		},
		chartPartPath,
		target,
		writes,
		{
			atomic,
			report: (issue) =>
				deps.reportWarning({
					code: issue.code,
					message: issue.message,
					severity: issue.severity,
					scope: 'save',
					slideId: slidePath,
					xmlPath: issue.partName,
				}),
		},
	);
	return written ? { complete, written } : { complete };
}
