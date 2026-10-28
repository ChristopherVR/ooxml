import { resolveManualPlotRect, type ChartFrameSize } from '../manual-layout';
import type { ChartViewModel } from './chart-view';
import type { Rect } from './chart-svg-util';

/** Share manual layout semantics while the painter supplies its measured axis bands. */
export function chartPlotRect(
	model: ChartViewModel,
	auto: Rect,
	outer: Rect,
	frame?: ChartFrameSize,
): Rect {
	if (!frame) return auto;
	const rect = resolveManualPlotRect(
		model.plotLayout,
		frame,
		{ x: auto.x, y: auto.y, width: auto.w, height: auto.h },
		{
			left: auto.x - outer.x,
			top: auto.y - outer.y,
			right: Math.max(0, outer.x + outer.w - auto.x - auto.w),
			bottom: Math.max(0, outer.y + outer.h - auto.y - auto.h),
		},
	);
	return rect ? { x: rect.x, y: rect.y, w: rect.width, h: rect.height } : auto;
}
