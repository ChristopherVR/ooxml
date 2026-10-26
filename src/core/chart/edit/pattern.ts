import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartSpace } from '../model';
import type { ChartFormatTarget } from './target';
import { chartTargetShape, setChartTargetFill } from './target-paint';

export interface ChartPatternEdit {
	preset?: string;
	foreground?: DrawingColor;
	background?: DrawingColor;
}

/** A partial pattern edit leaves unedited authored colours, transforms and source XML intact. */
export function setChartTargetPattern(
	space: ChartSpace,
	target: ChartFormatTarget,
	edit: ChartPatternEdit,
): ChartSpace {
	if (edit.preset !== undefined && !/^[A-Za-z][A-Za-z0-9]*$/u.test(edit.preset))
		throw new RangeError('Invalid chart pattern preset');
	const current = chartTargetShape(space, target)?.fill;
	const base: Extract<DrawingFill, { kind: 'pattern' }> =
		current?.kind === 'pattern'
			? structuredClone(current)
			: {
					kind: 'pattern',
					preset: 'pct5',
					foreground: { kind: 'srgb', value: '000000', transforms: [] },
					background: { kind: 'srgb', value: 'FFFFFF', transforms: [] },
				};
	return setChartTargetFill(space, target, { ...base, ...structuredClone(edit) });
}
