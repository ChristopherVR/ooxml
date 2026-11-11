// Where a data label can sit (`c:dLblPos`, ST_DLblPos) in each chart group. The schema lets every
// `c:dLbls` and `c:dLbl` carry any position, but Office accepts one only where its Format Data
// Labels pane offers it, and asks to repair a part with another (a line chart's `outEnd`): bars
// take the centre, inside end, inside base and (unstacked) outside end; lines, stock, scatter and
// bubble charts the centre and the four sides; pies and pies of pie the centre, inside end,
// outside end and best fit; 3-D bars and lines, areas, radars, doughnuts and surfaces none.
import type { ChartPlotGroup } from './model';

const SIDES = ['ctr', 'l', 'r', 't', 'b'] as const;
const PIE = ['ctr', 'inEnd', 'outEnd', 'bestFit'] as const;
const BAR = ['ctr', 'inEnd', 'inBase', 'outEnd'] as const;
const STACKED_BAR = ['ctr', 'inEnd', 'inBase'] as const;

/** The label positions a group of `element` with `grouping` accepts; empty when it takes none. */
export function chartLabelPositions(element: string, grouping?: string): readonly string[] {
	switch (element) {
		case 'barChart':
			return grouping === 'stacked' || grouping === 'percentStacked' ? STACKED_BAR : BAR;
		case 'lineChart':
		case 'stockChart':
		case 'scatterChart':
		case 'bubbleChart':
			return SIDES;
		case 'pieChart':
		case 'pie3DChart':
		case 'ofPieChart':
			return PIE;
		default:
			return [];
	}
}

/** Throws unless `group` accepts labels at `position`. */
export function assertChartLabelPosition(
	group: Pick<ChartPlotGroup, 'element' | 'grouping'>,
	position: string,
): void {
	if (!chartLabelPositions(group.element, group.grouping).includes(position))
		throw new RangeError(`A ${group.element} has no data label position ${position}`);
}
