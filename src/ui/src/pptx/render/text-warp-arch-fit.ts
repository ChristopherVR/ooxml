/**
 * How PowerPoint fits a line of text to the arc of Arch Up, Arch Down, Button
 * and Circle: text longer than the arc is not clipped or spilled over its ends,
 * it is scaled down (both ways, so glyph height follows) until it exactly fills
 * the arc. The arc is not always the whole half-ellipse (or loop): a start
 * angle close to the one where the text flips to the other half leaves only a
 * short stretch of it.
 *
 * Measured on PowerPoint 16 (rows of full-block glyphs and capital letters):
 *
 * - At 48 pt and at 60 pt the same capital row draws identically, as it must
 *   when the text length is pinned to the arc length.
 * - Arch Up, Arch Down and Button: the available sweep is `540 - 2a` degrees
 *   for a start angle `a` (turned half a turn when the text sits on the lower
 *   half, see `archOnUpperHalf`) between 180 and 270 degrees, the whole half
 *   ellipse before that. The capacity at 66, 70, 75, 80 and 84 degrees was
 *   168, 138, 102, 68 and 40 px for a 392 px box: the arc length of that sweep
 *   about the top or bottom of the box ellipse.
 * - Circle: the sweep about the tip is `2 (360 - a)` degrees from 180 to 360
 *   and `2 (180 - a)` from 0 to 180, the whole loop at the other end.
 */
import type { PptxTextWarpPreset, TextSegment, TextStyle } from 'ooxml-core/pptx';

import {
	archEllipseRadii,
	circleLoopRadii,
	normalizeWarpAngle,
	WARP_ARC_LINE_PITCH_EM,
	warpLineOutward,
} from './text-warp-arch';
import { DEFAULT_TEXT_FONT_SIZE } from '../constants';
import { measureTextAdvance } from './text-warp-envelope-measure';
import { resolveWarpRunPaint } from './text-warp-run-paint';
import { buildWarpPath } from './text-warp';

const RAD = Math.PI / 180;
const STEPS = 180;

/** Length of the ellipse arc `radiusAt(psi)` traces for `psi` in `[-half, half]` (radians). */
function arcLength(speed: (psi: number) => number, half: number): number {
	let length = 0;
	let previous = speed(-half);
	for (let i = 1; i <= STEPS; i++) {
		const current = speed(-half + (2 * half * i) / STEPS);
		length += ((previous + current) / 2) * ((2 * half) / STEPS);
		previous = current;
	}
	return length;
}

/** Sweep (degrees) of the arch ellipse left for the text: 180 is the whole half. */
export function archFitSweep(adj: number | undefined, half: 'up' | 'down'): number {
	const degrees = normalizeWarpAngle(adj ?? (half === 'up' ? 180 * 60000 : 0));
	const onUpperHalf = degrees >= 90 && degrees < 270;
	const angle = onUpperHalf ? degrees : (degrees + 180) % 360;
	return angle <= 180 ? 180 : Math.max(0, Math.min(180, 540 - 2 * angle));
}

/** Sweep (degrees) of the Circle loop left for the text: 360 is the whole loop. */
export function circleFitSweep(adj: number | undefined): number {
	const degrees = normalizeWarpAngle(adj ?? 180 * 60000);
	return Math.min(360, 2 * (degrees >= 180 ? 360 - degrees : 180 - degrees));
}

/**
 * The length (px) of the arc a warped line of `preset` may fill, or
 * `undefined` for a preset that does not scale its text to the path.
 */
export function warpFitCapacity(
	preset: string,
	width: number,
	height: number,
	outward: number,
	adj: number | undefined,
): number | undefined {
	if (preset === 'textArchUp' || preset === 'textButton' || preset === 'textArchDown') {
		const half = preset === 'textArchDown' ? 'down' : 'up';
		const { rx, ry } = archEllipseRadii(width, height, outward);
		// About the apex, `psi` from the apex: x = rx sin psi, y = ry cos psi.
		return arcLength(
			(psi) => Math.hypot(rx * Math.cos(psi), ry * Math.sin(psi)),
			(archFitSweep(adj, half) * RAD) / 2,
		);
	}
	if (preset === 'textCircle') {
		const { rx, ry } = circleLoopRadii(width, height, outward);
		// About the tip, `psi` from the tip: x = rx cos psi, y = ry sin psi.
		return arcLength(
			(psi) => Math.hypot(rx * Math.sin(psi), ry * Math.cos(psi)),
			(circleFitSweep(adj) * RAD) / 2,
		);
	}
	return undefined;
}

/** The scale (at most 1) that fits `textWidth` into `capacity`. */
export function warpFitScale(capacity: number | undefined, textWidth: number | undefined): number {
	if (!capacity || !textWidth || textWidth <= capacity) {
		return 1;
	}
	return Math.max(0.01, capacity / textWidth);
}

/** The advance width (px) of a paragraph's runs as rendered, or `undefined` with nothing to measure with. */
export function measureWarpParagraphWidth(
	segments: readonly TextSegment[],
	elementStyle: TextStyle | undefined,
): number | undefined {
	let total = 0;
	for (const segment of segments) {
		const paint = resolveWarpRunPaint(segment, elementStyle, '#000000', DEFAULT_TEXT_FONT_SIZE);
		const weight = paint.fontWeight;
		const width = measureTextAdvance(segment.text, {
			fontFamily: paint.fontFamily,
			fontSizePx: paint.fontSizePx,
			bold: weight === 'bold' || (weight !== undefined && Number(weight) >= 600),
			italic: paint.fontStyle === 'italic',
		});
		if (width === undefined) {
			return undefined;
		}
		total += width;
	}
	return total;
}

/** One paragraph's warp placement: its SVG path and the factor its font sizes are multiplied by. */
export interface WarpParagraphLayout {
	d: string;
	/** 1 unless the text is longer than the arc (see the module header). */
	fitScale: number;
}

/**
 * The path and the fit scale of every paragraph of a warped block, the one
 * call every binding makes for the presets drawn along a path. Lines of the
 * arch and loop presets lie one line pitch ({@link WARP_ARC_LINE_PITCH_EM}
 * of the block's largest font) apart.
 */
export function warpParagraphLayouts(
	preset: PptxTextWarpPreset,
	width: number,
	height: number,
	paragraphs: readonly { segments: readonly TextSegment[] }[],
	adj: number | undefined,
	adj2: number | undefined,
	elementStyle: TextStyle | undefined,
): WarpParagraphLayout[] {
	let largest = 0;
	for (const paragraph of paragraphs) {
		for (const segment of paragraph.segments) {
			largest = Math.max(
				largest,
				resolveWarpRunPaint(segment, elementStyle, '#000000', DEFAULT_TEXT_FONT_SIZE).fontSizePx,
			);
		}
	}
	const gap = largest * WARP_ARC_LINE_PITCH_EM;
	const count = paragraphs.length;
	return paragraphs.map((paragraph, index) => {
		const capacity = warpFitCapacity(
			preset,
			width,
			height,
			warpLineOutward(preset, index, count, gap),
			adj,
		);
		return {
			d: buildWarpPath(preset, width, height, index, count, adj, adj2, gap),
			fitScale:
				capacity === undefined
					? 1
					: warpFitScale(capacity, measureWarpParagraphWidth(paragraph.segments, elementStyle)),
		};
	});
}
