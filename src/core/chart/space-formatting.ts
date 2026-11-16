// The direct formatting (`readChartFormatting`) of the part a `ChartSpace` writes, which the painter
// reads for the chart area, plot area, title and legend paint. A host that paints an edited model
// before saving it reads the formatting this way, so the preview shows what will be saved.
import { parseXml } from '../xml/index';
import type { ChartSpace } from './model';
import { readChartFormatting } from './read-formatting';
import type { ChartStyleDefinition } from './style-definition';
import { writeChartSpace } from './write-space';

/** The direct formatting `space` saves, or undefined for a chart with no chart group. */
export function chartSpaceFormatting(space: ChartSpace): ChartStyleDefinition | undefined {
	if (!space.plotArea.groups.length) return undefined;
	const root = parseXml(writeChartSpace(space), { label: 'Chart part' }).documentElement;
	return root ? readChartFormatting(root) : undefined;
}
