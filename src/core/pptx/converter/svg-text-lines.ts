import type { TextSegment } from '../core/types/text';
import {
	DEFAULT_FONT_ADVANCE_TABLE,
	FONT_ADVANCE_TABLES,
} from '../../text/font-metrics/font-advance-widths.generated';
import { measureTextWidth } from '../../diagram/layout/smartart-text-wrap-fit';
import { wrapStyledRuns } from '../../text/wrap-styled-runs';

export interface SvgTextRun {
	text: string;
	style: TextSegment['style'];
	lineStart: boolean;
}
export type SvgTextMeasurer = (
	text: string,
	font: { family: string; size: number; bold: boolean; italic: boolean },
) => number;

/** Basic horizontal wrapping for static exports. Unknown fonts use fallback metrics. */
export function svgTextLines(
	segments: TextSegment[],
	width: number,
	fontSize: number,
	fontFamily: string,
	measure?: SvgTextMeasurer,
): SvgTextRun[] {
	const lines = wrapStyledRuns(
		segments.map((segment) => ({
			...segment,
			text: segment.isParagraphBreak ? '\n' : segment.text,
		})),
		width,
		(content, segment) => {
			const size = segment.style.fontSize ?? fontSize;
			const family = segment.style.fontFamily ?? fontFamily;
			const table = FONT_ADVANCE_TABLES[family] ?? DEFAULT_FONT_ADVANCE_TABLE;
			const measured = measure?.(content, {
				family,
				size,
				bold: !!segment.style.bold,
				italic: !!segment.style.italic,
			});
			return measured !== undefined && Number.isFinite(measured) && measured >= 0
				? measured
				: measureTextWidth(content, size, table);
		},
	);
	return lines.flatMap((line) =>
		line.map((run, index) => ({ text: run.text, style: run.style, lineStart: index === 0 })),
	);
}
