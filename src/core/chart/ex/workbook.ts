// Structural ChartEx data uses the shared chart workbook allocator and atomic source validator.
import { chartExDataLayout } from './data-layout';
import { chartExSpaceWorkbookWrites } from './workbook-writes';
import type { ChartExSpace } from './model';
import { appendChartWorkbookSheet } from '../workbook/append-sheet';
import { writeChartWorkbookUpdates } from '../workbook/writer';
import { openEmbeddedPackage } from '../workbook/writer';
import { NS, children, first, parseXml } from '../../xml';

/** Validate the old literal sources, then append a dedicated sheet for expanded chart data. */
export async function appendChartExWorkbookData(
	bytes: Uint8Array,
	before: ChartExSpace,
	after: ChartExSpace,
) {
	const zip = await openEmbeddedPackage(bytes);
	const xml = await zip.file('xl/workbook.xml')?.async('string');
	if (!xml) throw new Error('The chart workbook is missing.');
	const root = parseXml(xml).documentElement;
	const defined = first(root, 'definedNames', NS.x);
	const names = defined ? children(defined, 'definedName', NS.x) : [];
	const resolve = (name: string) =>
		names.find(
			(entry) =>
				entry.getAttribute('name')?.toLowerCase() === name.toLowerCase() &&
				!entry.hasAttribute('localSheetId'),
		)?.textContent ?? undefined;
	await writeChartWorkbookUpdates(bytes, chartExSpaceWorkbookWrites(before, resolve), {
		atomic: true,
		requireCurrentValues: true,
	});
	let layout: ReturnType<typeof chartExDataLayout> | undefined;
	const result = await appendChartWorkbookSheet(
		bytes,
		(name) => {
			layout = chartExDataLayout(after, name);
			return layout.rows;
		},
		{ nameBase: 'ChartExData' },
	);
	return { bytes: result.bytes, space: layout!.space };
}
