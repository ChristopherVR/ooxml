// Source-backed ChartEx writes refuse parser losses and retain the producer's XML write options.
import { sameChartModel } from '../keep-source';
import { CHART_POINTS_LIMITED } from '../point-limit';
import type { ChartExSpace } from './model';
import { parseChartExSpace } from './parse';
import { chartExWriteOptionsOf, writeChartExSpace } from './write';

/** A loaded ChartEx part is editable only when no parser issue would be lost. */
export function writeChartExSpaceOverSource(source: string, space: ChartExSpace): string {
	const parsed = parseChartExSpace(source);
	if (sameChartModel(parsed.chartExSpace, space)) return source;
	if (parsed.issues.some((issue) => issue.code !== CHART_POINTS_LIMITED))
		throw new Error(
			'This extended chart contains content that cannot be preserved through an edit.',
		);
	return writeChartExSpace(space, chartExWriteOptionsOf(source));
}
