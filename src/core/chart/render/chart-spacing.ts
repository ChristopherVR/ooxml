import { DEFAULT_BAR_GAP_WIDTH } from '../bar-cluster-geometry';
import type { ChartSummary } from './summary';

type ChartSpacing = Pick<ChartSummary, 'barGapWidth' | 'barOverlap' | 'grouping'>;

/** Effective authored spacing. Imported charts already carry OOXML's explicit defaults. */
export function chartBarSpacing(
	chart: ChartSpacing,
): Required<Pick<ChartSummary, 'barGapWidth' | 'barOverlap'>> {
	const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
	return {
		barGapWidth: chart.barGapWidth ?? DEFAULT_BAR_GAP_WIDTH,
		barOverlap: chart.barOverlap ?? (stacked ? 100 : 0),
	};
}
