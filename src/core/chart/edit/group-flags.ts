// Chart group settings edited on their own, outside a type change: whether a group varies its
// point colours (`c:varyColors`) and the direction of a bar chart (`c:barDir`). The scatter style
// and the grouping have their own operations (`group-options.ts`, `type.ts`). Values follow
// ECMA-376 Part 1, 21.2.2 (CT_Boolean, ST_BarDir).
import type { ChartSpace } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { GROUP_FIELDS } from './type-schema';
import { rotatedAxisPosition } from './type-axes';

/**
 * Sets or removes (`null`) `c:varyColors` of the group holding series `seriesIndex`. A group
 * element without the setting (stock, surface) is refused rather than written into a part Office
 * would repair.
 */
export function setChartVaryColors(
	space: ChartSpace,
	seriesIndex: number,
	vary: boolean | null,
): ChartSpace {
	if (vary !== null && typeof vary !== 'boolean')
		throw new RangeError('Chart varyColors is a flag');
	return editChartSpace(space, (draft) => {
		const { group } = chartSeriesEntryAt(draft, seriesIndex);
		if (!(GROUP_FIELDS[group.element] ?? []).includes('varyColors'))
			throw new RangeError(`A ${group.element} has no varyColors setting`);
		if (vary === null) delete group.varyColors;
		else group.varyColors = vary;
	});
}

/**
 * Turns a bar chart between vertical columns (`col`) and horizontal bars (`bar`). Every group is a
 * bar group, as a column chart becomes a bar chart in Office's Change Chart Type: the axes of the
 * groups swap sides (bottom and left, top and right) and keep everything else. A combination with
 * other families is refused; one series changes direction through `setChartSeriesType`.
 */
export function setChartBarDirection(space: ChartSpace, direction: 'bar' | 'col'): ChartSpace {
	if (direction !== 'bar' && direction !== 'col')
		throw new RangeError(`Invalid chart bar direction: ${String(direction)}`);
	const groups = space.plotArea.groups;
	if (groups.length === 0 || groups.some((group) => group.kind !== 'bar'))
		throw new RangeError('Only a chart of bar groups changes its bar direction');
	return editChartSpace(space, (draft) => {
		const rotated = new Set<number>();
		for (const group of draft.plotArea.groups) {
			if ((group.barDirection ?? 'col') === direction) continue;
			group.barDirection = direction;
			for (const id of group.axisIds) rotated.add(id);
		}
		for (const axis of draft.plotArea.axes) {
			if (axis.id === undefined || !rotated.has(axis.id)) continue;
			const position = rotatedAxisPosition(axis.position);
			if (position === undefined) delete axis.position;
			else axis.position = position;
		}
	});
}
