/**
 * The markup area of the Word PDF export, laid out the way Word prints "Document showing markup"
 * with balloons: the page is shrunk so that a markup pane fits beside its text, and each balloon
 * (a comment, a formatting change or, when every revision is shown in balloons, a deletion) sits
 * in the pane as near its anchor's line as the balloons above it allow, joined to the anchor by
 * a dotted connector in its author's colour. Balloon text is real PDF text. Coordinates are the
 * page's CSS pixels, y down, before the shrink. DOM-free.
 */
import type {
	LayoutFragment,
	LayoutLine,
	LayoutPageBox,
	TextMeasurer,
} from 'ooxml-core/docx/layout';
import type { PdfVectorPage } from '../../export/pdf-vector';
import { cssRgba, type DocxPdfText } from './pdf-text';

const PT_PER_PX = 0.75;
/** Word's default balloon width preference (3.7"), the room the pane takes beside the text. */
export const BALLOON_PANE_PX = 3.7 * 96;
/** Gap between the right text margin and the pane, and pane insets, as Word prints them. */
const PANE_GAP_PX = 12;
const BALLOON_INSET_LEFT_PX = 30;
const BALLOON_INSET_RIGHT_PX = 6;
const BALLOON_PADDING_X_PX = 6;
const BALLOON_PADDING_Y_PX = 3;
const BALLOON_GAP_PX = 2;
const BALLOON_FONT_PT = 9;
/** The font Word writes balloon labels in. */
export const LABEL_FAMILY = 'Segoe UI';
const PANE_FILL = '#f2f2f2';

/** How the shrunk page sits on the sheet: the scale and the vertical offset, in points. */
export interface DocxPdfBalloonFrame {
	scale: number;
	offsetYPt: number;
	/** The `cm` matrix that shrinks the page (applied before the page's own transform). */
	matrix: [number, number, number, number, number, number];
}

/** The frame of a page shown with a markup pane (Word's print layout for balloons). */
export function balloonFrame(
	page: Pick<LayoutPageBox, 'widthPx' | 'heightPx' | 'marginRightPx'>,
): DocxPdfBalloonFrame {
	const scale = page.widthPx / (page.widthPx - page.marginRightPx + BALLOON_PANE_PX);
	const offsetYPt = ((1 - scale) * page.heightPx * PT_PER_PX) / 2;
	return { scale, offsetYPt, matrix: [scale, 0, 0, scale, 0, offsetYPt] };
}

/** A run of balloon text in one font. */
export interface BalloonPiece {
	text: string;
	bold?: boolean;
	family?: string;
	color?: string;
}

export interface Balloon {
	kind: 'comment' | 'format' | 'delete';
	/** The author's colour (border and connector). */
	color: string;
	/** Balloon fill; white when absent. */
	fill?: string;
	/** Paragraphs of the balloon, each a list of pieces. */
	paragraphs: BalloonPiece[][];
	/** The anchor: where the connector starts (x and the line's bottom) and the line's top. */
	anchorX: number;
	anchorTop: number;
	anchorBottom: number;
	/** Draw a small marker at the anchor (revisions; comments highlight their range instead). */
	marker: boolean;
	/** Placed balloon box, set by {@link layoutBalloons}. */
	box?: { x: number; y: number; width: number; height: number };
	/** Called with the placed box, e.g. to move the comment's PDF note onto it. */
	placed?: (box: NonNullable<Balloon['box']>) => void;
}

interface PlacedPiece {
	piece: BalloonPiece;
	text: string;
	x: number;
	width: number;
}

const fontSize = () => (BALLOON_FONT_PT * 96) / 72;
const spec = (piece: BalloonPiece) => ({
	family: piece.family ?? LABEL_FAMILY,
	sizePx: fontSize(),
	...(piece.bold ? { bold: true } : {}),
});

/** Greedy word wrap of a balloon's paragraphs into lines no wider than `width`. */
export function wrapBalloon(
	paragraphs: readonly BalloonPiece[][],
	width: number,
	measurer: TextMeasurer,
): PlacedPiece[][] {
	const lines: PlacedPiece[][] = [];
	for (const paragraph of paragraphs) {
		let line: PlacedPiece[] = [];
		let x = 0;
		const push = (piece: BalloonPiece, token: string) => {
			let text = token;
			if (x > 0 && x + measurer.widthOf(text.trimEnd(), spec(piece)) > width) {
				lines.push(line);
				line = [];
				x = 0;
				text = text.trimStart();
				if (!text) return;
			}
			const advance = measurer.widthOf(text, spec(piece));
			const last = line.at(-1);
			if (last && last.piece === piece) {
				last.text += text;
				last.width += advance;
			} else line.push({ piece, text, x, width: advance });
			x += advance;
		};
		for (const piece of paragraph)
			for (const token of piece.text.match(/\n|[^\s]+\s*|[^\S\n]+/g) ?? []) {
				if (token === '\n') {
					lines.push(line);
					line = [];
					x = 0;
				} else push(piece, token);
			}
		lines.push(line);
	}
	return lines;
}

const lineHeight = (measurer: TextMeasurer) =>
	measurer.lineHeightOf({ family: LABEL_FAMILY, sizePx: fontSize() });

/** The pane's horizontal extent and the balloons' (page pixels before the shrink). */
export function paneGeometry(page: Pick<LayoutPageBox, 'widthPx' | 'heightPx' | 'marginRightPx'>) {
	const { scale } = balloonFrame(page);
	const left = page.widthPx - page.marginRightPx + PANE_GAP_PX;
	const right = page.widthPx / scale;
	return {
		left,
		right,
		balloonLeft: left + BALLOON_INSET_LEFT_PX,
		balloonWidth: right - BALLOON_INSET_RIGHT_PX - left - BALLOON_INSET_LEFT_PX,
	};
}

/**
 * Places the balloons in the pane: in anchor order, each at its anchor line's top unless the
 * balloon above pushes it down; balloons running off the page are pushed back up. Returns false
 * when they do not all fit (some then overlap at the top).
 */
export function layoutBalloons(
	balloons: Balloon[],
	page: Pick<LayoutPageBox, 'widthPx' | 'heightPx' | 'marginRightPx'>,
	measurer: TextMeasurer,
): { fits: boolean; lines: Map<Balloon, PlacedPiece[][]> } {
	const { balloonLeft, balloonWidth } = paneGeometry(page);
	const order = [...balloons].sort((a, b) => a.anchorTop - b.anchorTop || a.anchorX - b.anchorX);
	const lines = new Map<Balloon, PlacedPiece[][]>();
	let bottom = -Infinity;
	for (const balloon of order) {
		const wrapped = wrapBalloon(
			balloon.paragraphs,
			balloonWidth - 2 * BALLOON_PADDING_X_PX,
			measurer,
		);
		lines.set(balloon, wrapped);
		const height = wrapped.length * lineHeight(measurer) + 2 * BALLOON_PADDING_Y_PX;
		const y = Math.max(balloon.anchorTop, bottom + BALLOON_GAP_PX);
		balloon.box = { x: balloonLeft, y, width: balloonWidth, height };
		bottom = y + height;
	}
	let limit = page.heightPx - BALLOON_GAP_PX;
	for (const balloon of [...order].reverse()) {
		const box = balloon.box!;
		if (box.y + box.height <= limit) break;
		box.y = limit - box.height;
		limit = box.y - BALLOON_GAP_PX;
	}
	const fits = !order.length || order[0]!.box!.y >= 0;
	if (!fits) for (const balloon of order) balloon.box!.y = Math.max(0, balloon.box!.y);
	for (const balloon of order) balloon.placed?.(balloon.box!);
	return { fits, lines };
}

const n = (value: number) => Number(value.toFixed(3));

function rgb(page: PdfVectorPage, color: string, stroke = false): boolean {
	const rgba = cssRgba(color);
	if (!rgba || rgba.a === 0) return false;
	if (stroke) page.strokeColor(rgba.r, rgba.g, rgba.b);
	else page.fillColor(rgba.r, rgba.g, rgba.b);
	return true;
}

/** A rounded rectangle path. */
function roundedRect(page: PdfVectorPage, x: number, y: number, w: number, h: number, r: number) {
	const k = r * 0.5523;
	page.op(
		[
			`${n(x + r)} ${n(y)} m`,
			`${n(x + w - r)} ${n(y)} l`,
			`${n(x + w - r + k)} ${n(y)} ${n(x + w)} ${n(y + r - k)} ${n(x + w)} ${n(y + r)} c`,
			`${n(x + w)} ${n(y + h - r)} l`,
			`${n(x + w)} ${n(y + h - r + k)} ${n(x + w - r + k)} ${n(y + h)} ${n(x + w - r)} ${n(y + h)} c`,
			`${n(x + r)} ${n(y + h)} l`,
			`${n(x + r - k)} ${n(y + h)} ${n(x)} ${n(y + h - r + k)} ${n(x)} ${n(y + h - r)} c`,
			`${n(x)} ${n(y + r)} l`,
			`${n(x)} ${n(y + r - k)} ${n(x + r - k)} ${n(y)} ${n(x + r)} ${n(y)} c h`,
		].join(' '),
	);
}

/** Paints the pane, the connectors and the balloons with their text. */
export function paintBalloons(
	page: PdfVectorPage,
	layout: Pick<LayoutPageBox, 'widthPx' | 'heightPx' | 'marginRightPx'>,
	balloons: Balloon[],
	text: DocxPdfText,
): boolean {
	const geometry = paneGeometry(layout);
	const { fits, lines } = layoutBalloons(balloons, layout, text.measurer);
	rgb(page, PANE_FILL);
	page.rect(geometry.left, 0, geometry.right - geometry.left, layout.heightPx);
	page.op('f');
	const height = lineHeight(text.measurer);
	for (const balloon of balloons) {
		const box = balloon.box!;
		// The connector: along the anchor line's bottom to the pane, then to the balloon.
		if (rgb(page, balloon.color, true)) {
			page.lineStyle(0.6, 0, 0, [1, 1]);
			const elbow = geometry.left - 6;
			page.op(
				`${n(balloon.anchorX)} ${n(balloon.anchorBottom)} m ${n(elbow)} ${n(balloon.anchorBottom)} l ${n(box.x)} ${n(box.y + 8)} l S`,
			);
		}
		if (balloon.marker && rgb(page, balloon.color)) {
			const { anchorX: x, anchorBottom: y } = balloon;
			page.op(`${n(x - 2)} ${n(y + 1)} m ${n(x + 2)} ${n(y + 1)} l ${n(x)} ${n(y - 3)} l h f`);
		}
		rgb(page, balloon.fill ?? '#ffffff');
		rgb(page, balloon.color, true);
		page.lineStyle(0.6, 0, 0);
		roundedRect(page, box.x, box.y, box.width, box.height, 3);
		page.op('B');
		(lines.get(balloon) ?? []).forEach((pieces, index) => {
			const top = box.y + BALLOON_PADDING_Y_PX + index * height;
			const line: LayoutLine = {
				yPx: 0,
				heightPx: height,
				baselinePx: height * 0.8,
				fragments: [],
				sourceStart: 0,
				sourceEnd: 0,
			};
			for (const placed of pieces) {
				if (!placed.text.trim()) continue;
				const fragment: LayoutFragment = {
					text: placed.text,
					xPx: placed.x,
					widthPx: placed.width,
					runIndex: 0,
					fontFamily: placed.piece.family ?? LABEL_FAMILY,
					fontSizePt: BALLOON_FONT_PT,
					color: placed.piece.color ?? '#000000',
					...(placed.piece.bold ? { bold: true } : {}),
				};
				text.draw(page, fragment, box.x + BALLOON_PADDING_X_PX + placed.x, line, top);
			}
		});
	}
	return fits;
}
