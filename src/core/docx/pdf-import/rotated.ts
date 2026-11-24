/**
 * Rotated text: spans off the horizontal are grouped by direction, read as lines and paragraphs in
 * their own frame (the ordinary line and paragraph rules, applied after turning them level), and
 * become text boxes turned like the original: text running up or down uses the box's vertical
 * text (`vert270`, `vert`), other angles a rotated box.
 */
import type { PdfTextSpan } from '../../pdf/index';
import { pageLines, type Line } from './lines';
import { buildParagraphs, type ParagraphBlock } from './paragraphs';

/** Spans off the horizontal by more than this (radians) are rotated text. */
export const HORIZONTAL = 0.05;

/** A text box at a place on the page. */
export interface FloatingText {
	page: number;
	/** The box as it sits on the page (reading coordinates, points). */
	x: number;
	y: number;
	width: number;
	height: number;
	/** The text direction in degrees, counter-clockwise (90: bottom to top). */
	angle: number;
	/** The box before its rotation (for angles other than 90 and 270): text length and height. */
	length: number;
	thickness: number;
	blocks: ParagraphBlock[];
}

const turn = (span: PdfTextSpan, angle: number): PdfTextSpan => {
	const cos = Math.cos(angle);
	const sin = Math.sin(angle);
	return { ...span, x: span.x * cos - span.y * sin, y: span.x * sin + span.y * cos, angle: 0 };
};
const unturn = (u: number, v: number, angle: number): [number, number] => {
	const cos = Math.cos(angle);
	const sin = Math.sin(angle);
	return [u * cos + v * sin, -u * sin + v * cos];
};

/** Groups the lines of one direction into boxes: lines close together that overlap along it. */
function boxes(lines: Line[]): Line[][] {
	const out: Line[][] = [];
	for (const line of lines) {
		const box = out.find((group) => {
			const last = group[group.length - 1]!;
			return (
				line.baseline - last.baseline < Math.max(last.size, line.size) * 2 &&
				line.left < Math.max(...group.map((item) => item.right)) + line.size &&
				line.right > Math.min(...group.map((item) => item.left)) - line.size
			);
		});
		if (box) box.push(line);
		else out.push([line]);
	}
	return out;
}

/** The page's rotated text as text boxes, and the number of lines they hold. */
export function rotatedText(
	page: number,
	spans: readonly PdfTextSpan[],
	bodySize: number,
): { floats: FloatingText[]; lines: number } {
	const byAngle = new Map<number, PdfTextSpan[]>();
	for (const span of spans) {
		if (Math.abs(span.angle) < HORIZONTAL || span.invisible || !span.text.length) continue;
		const degrees = (Math.round((span.angle * 180) / Math.PI) + 360) % 360;
		const group = byAngle.get(degrees) ?? [];
		group.push(span);
		byAngle.set(degrees, group);
	}
	const floats: FloatingText[] = [];
	let count = 0;
	for (const [degrees, group] of byAngle) {
		const angle = (degrees * Math.PI) / 180;
		const { lines } = pageLines(
			page,
			group.map((span) => turn(span, angle)),
		);
		for (const box of boxes(lines)) {
			count += box.length;
			const left = Math.min(...box.map((line) => line.left));
			const right = Math.max(...box.map((line) => line.right));
			const top = Math.min(...box.map((line) => line.top));
			const bottom = Math.max(...box.map((line) => line.bottom));
			const corners = [
				unturn(left, top, angle),
				unturn(right, top, angle),
				unturn(left, bottom, angle),
				unturn(right, bottom, angle),
			];
			const xs = corners.map((point) => point[0]);
			const ys = corners.map((point) => point[1]);
			const x = Math.min(...xs);
			const y = Math.min(...ys);
			const blocks = buildParagraphs(box, { left, right }, bodySize).map(
				({ heading: _heading, ...block }) => ({ ...block, spaceBefore: 0 }),
			);
			floats.push({
				page,
				x,
				y,
				width: Math.max(...xs) - x,
				height: Math.max(...ys) - y,
				angle: degrees,
				length: right - left,
				thickness: bottom - top,
				blocks,
			});
		}
	}
	return { floats, lines: count };
}
