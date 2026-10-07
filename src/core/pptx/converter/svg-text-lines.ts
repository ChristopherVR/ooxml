import type { TextSegment } from '../core/types/text';
import {
	DEFAULT_FONT_ADVANCE_TABLE,
	FONT_ADVANCE_TABLES,
} from '../core/utils/font-advance-widths.generated';
import { measureTextWidth } from '../core/utils/smartart-text-wrap-fit';

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
	const result: SvgTextRun[] = [];
	let used = 0;
	let lineStart = true;
	let breakable = false;
	for (const segment of segments) {
		if (segment.isParagraphBreak) {
			used = 0;
			lineStart = true;
			breakable = false;
			continue;
		}
		const size = segment.style.fontSize ?? fontSize;
		const family = segment.style.fontFamily ?? fontFamily;
		const table = FONT_ADVANCE_TABLES[family] ?? DEFAULT_FONT_ADVANCE_TABLE;
		let text = '';
		let startsLine = lineStart;
		const flush = () => {
			if (text) result.push({ text, style: segment.style, lineStart: startsLine });
			text = '';
			startsLine = lineStart;
		};
		for (const token of segment.text.split(/(\n|[^\S\n]+)/u).filter(Boolean)) {
			if (token === '\n') {
				flush();
				used = 0;
				lineStart = true;
				startsLine = true;
				breakable = false;
				continue;
			}
			const whitespace = /^\s+$/u.test(token);
			const measured = measure?.(token, {
				family,
				size,
				bold: !!segment.style.bold,
				italic: !!segment.style.italic,
			});
			const advance =
				measured !== undefined && Number.isFinite(measured) && measured >= 0
					? measured
					: measureTextWidth(token, size, table);
			if (!whitespace && breakable && used + advance > width) {
				text = text.trimEnd();
				flush();
				used = 0;
				lineStart = true;
				startsLine = true;
			}
			if (whitespace && lineStart) continue;
			text += token;
			used += advance;
			lineStart = false;
			breakable = whitespace;
		}
		flush();
	}
	return result;
}
