/** Small composed pieces the slides share: legends, stat blocks, cards and chart annotations. */
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import { FONTS } from '../../lib/fiction';
import { pt, type Box } from '../kit/geometry';
import type { Slide } from '../kit/slide';
import { C, label, rule } from './style';

export interface LegendItem {
	label: string;
	color: string;
}

/**
 * A row of legend chips (a colour dot and a name in a pill). Returns the pills, so a slide can
 * use one as the trigger of an interactive sequence.
 */
export function legend(slide: Slide, items: LegendItem[], x: number, y: number): PptxElement[] {
	let left = x;
	return items.map((item) => {
		const width = 40 + item.label.length * 7;
		const pill = slide.shape(`Legend ${item.label}`, 'roundRect', {
			x: left,
			y,
			width,
			height: 24,
			adjustments: { adj: 50000 },
			fill: { type: 'solid', color: C.white },
			stroke: { color: C.rule, width: 1 },
			text: item.label,
			textStyle: {
				fontSize: pt(9.5),
				fontFamily: FONTS.body,
				color: C.ink,
				alignment: 'left',
				verticalAlignment: 'middle',
			},
			insets: { left: 26, top: 0, right: 6, bottom: 0 },
			wrap: false,
		});
		slide.shape(`Legend dot ${item.label}`, 'ellipse', {
			x: left + 9,
			y: y + 7,
			width: 10,
			height: 10,
			fill: { type: 'solid', color: item.color },
			stroke: { color: 'transparent' },
		});
		left += width + 8;
		return pill;
	});
}

/** A figure with a rule above and a caption below, as used in the side columns. */
export function stat(
	slide: Slide,
	name: string,
	figure: string,
	caption: string,
	box: Box,
	color: string = C.ink,
): PptxElement[] {
	return [
		rule(slide, `${name} rule`, { x: box.x, y: box.y, width: box.width, height: 2 }, C.ink),
		label(
			slide,
			`${name} figure`,
			figure,
			{ x: box.x, y: box.y + 14, width: box.width, height: 44 },
			{ size: 30, display: true, bold: true, color, tracking: -0.5 },
		),
		label(
			slide,
			`${name} caption`,
			caption,
			{ x: box.x, y: box.y + 62, width: box.width, height: box.height - 62 },
			{ size: 11, color: C.slate, leading: 1.3 },
		),
	];
}

/** A white card with a soft shadow. */
export function card(slide: Slide, name: string, box: Box): PptxElement {
	return slide.shape(name, 'roundRect', {
		...box,
		adjustments: { adj: 3200 },
		fill: { type: 'solid', color: C.white },
		stroke: { color: '#EFEADF', width: 1 },
		shadow: { color: '#1B2430', blur: 28, offsetX: 0, offsetY: 10, opacity: 0.1 },
	});
}

/**
 * A chart annotation: a short text and a thin leader line from the text to the point it is
 * about. The leader is a custom-geometry elbow, so it can turn a corner around the data.
 */
export function annotation(
	slide: Slide,
	name: string,
	text: string,
	box: Box,
	target: { x: number; y: number },
	align: 'left' | 'right' = 'right',
): PptxElement[] {
	const note = label(slide, `${name} note`, text, box, {
		size: 11,
		color: C.coral,
		bold: true,
		align,
		leading: 1.25,
	});
	// From the middle of the text's near edge, across, then down (or up) to the target.
	const startX = align === 'right' ? box.x + box.width + 8 : box.x - 8;
	const startY = box.y + 12;
	const x = Math.min(startX, target.x);
	const y = Math.min(startY, target.y);
	const width = Math.max(1, Math.abs(target.x - startX));
	const height = Math.max(1, Math.abs(target.y - startY));
	const fromLeft = startX <= target.x;
	const fromTop = startY <= target.y;
	const path = `M ${fromLeft ? 0 : width} ${fromTop ? 0 : height} L ${fromLeft ? width : 0} ${fromTop ? 0 : height} L ${fromLeft ? width : 0} ${fromTop ? height : 0}`;
	const leader = slide.freeform(`${name} leader`, path, {
		x,
		y,
		width,
		height,
		fill: { type: 'none' },
		stroke: { color: C.coral, width: 1.25 },
	});
	const dot = slide.shape(`${name} dot`, 'ellipse', {
		x: target.x - 4,
		y: target.y - 4,
		width: 8,
		height: 8,
		fill: { type: 'solid', color: C.coral },
		stroke: { color: C.paper, width: 1.5 },
	});
	return [note, leader, dot];
}
