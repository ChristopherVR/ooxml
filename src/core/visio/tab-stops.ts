import type { VisioTabStop } from './shape-sections-types';

/** Visio's default tab spacing in inches, used past the last explicit stop (and with none). */
export const VISIO_DEFAULT_TAB_SPACING = 0.5;

/** The text after one tab character (the first segment is the text before any tab). */
export interface VisioTabSegment {
	/** Advance width of the segment's text in inches. */
	width: number;
	/** Width of the text before its first decimal point, or the whole width when it has none. */
	decimal: number;
}

/**
 * Where each segment of a tabbed line starts. `origin` is the left edge tab positions are measured
 * from and `start` where the line's first segment begins (both on the same axis). A tab advances
 * to the next stop to the right of the text so far: an explicit stop of the paragraph, then the
 * default spacing. Left stops start the text there, centre stops centre it, right stops end it
 * there and decimal stops put the decimal point there; text that would start before the end of the
 * previous segment follows it directly.
 */
export function visioTabStarts(
	segments: readonly VisioTabSegment[],
	tabs: readonly VisioTabStop[] = [],
	origin = 0,
	start = 0,
	spacing = VISIO_DEFAULT_TAB_SPACING,
): number[] {
	const starts: number[] = [];
	let cursor = start;
	segments.forEach((segment, index) => {
		let x = start;
		if (index > 0) {
			const edge = cursor + 1e-9;
			const explicit = tabs.find((stop) => origin + stop.position > edge);
			const stop = explicit ?? {
				position: (Math.floor((cursor - origin) / spacing + 1e-9) + 1) * spacing,
				align: 'left' as const,
			};
			const at = origin + stop.position;
			x =
				stop.align === 'center'
					? at - segment.width / 2
					: stop.align === 'right'
						? at - segment.width
						: stop.align === 'decimal'
							? at - segment.decimal
							: at;
			x = Math.max(x, cursor);
		}
		starts.push(x);
		cursor = x + segment.width;
	});
	return starts;
}
