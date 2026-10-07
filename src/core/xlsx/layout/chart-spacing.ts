import type { ChartObject } from '../model';

type ChartSpacing = Pick<ChartObject, 'barGapWidth' | 'barOverlap' | 'grouping'>;

/** Effective authored spacing. Imported charts already carry OOXML's explicit defaults. */
export function chartBarSpacing(
	chart: ChartSpacing,
): Required<Pick<ChartObject, 'barGapWidth' | 'barOverlap'>> {
	const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
	return {
		barGapWidth: chart.barGapWidth ?? 150,
		barOverlap: chart.barOverlap ?? (stacked ? 100 : 0),
	};
}
