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
import type { CompatibilityWarningInput } from '../../services/PptxCompatibilityService';
import type { PptxExternalData } from '../../types';
import type { PptxChartWorkbookWrite } from '../../utils/chart-xlsx-writer';

/** The runtime primitives this module needs, kept explicit for testability. */
export interface ChartExternalWorkbookSaveDeps {
	zip: JSZip;
	resolveImagePath(basePath: string, target: string): string;
	reportWarning(warning: CompatibilityWarningInput): void;
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
): Promise<boolean> {
	const target = externalData
		? {
				...(externalData.targetPath
					? { partName: deps.resolveImagePath(chartPartPath, externalData.targetPath) }
					: {}),
				...(externalData.isExternal ? { external: true } : {}),
			}
		: undefined;
	return saveChartWorkbookUpdates(
		{
			read: async (partName) => deps.zip.file(partName)?.async('uint8array'),
			write: (partName, bytes) => {
				deps.zip.file(partName, bytes);
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
}
