// Source-backed ChartEx writes refuse parser losses and retain the producer's XML write options.
import type { ChartExSpace } from './model';
import { parseChartExSpace } from './parse';
import { chartExWriteOptionsOf, writeChartExSpace } from './write';

/** A loaded ChartEx part is editable only when no parser issue would be lost. */
export function writeChartExSpaceOverSource(source: string, space: ChartExSpace): string {
	const parsed = parseChartExSpace(source);
	if (JSON.stringify(parsed.chartExSpace) === JSON.stringify(space)) return source;
	if (parsed.issues.length)
		throw new Error(
			'This extended chart contains content that cannot be preserved through an edit.',
		);
	return writeChartExSpace(space, chartExWriteOptionsOf(source));
}
