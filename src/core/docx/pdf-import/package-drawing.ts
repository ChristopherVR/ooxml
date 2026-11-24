/**
 * Writes a vector drawing (`drawings.ts`) as a floating DrawingML object placed on the page where
 * the PDF drew it, behind the text: one `wps:wsp` shape (preset geometry when the piece is a
 * rectangle, ellipse or line, else `a:custGeom`), one custom-geometry shape with a path per piece
 * when all pieces look the same, or a `wpg:wgp` group of shapes otherwise.
 */
import type { PdfPathSegment } from '../../pdf/index';
import { NS } from '../load/legacy-doc/xml';
import type { DrawingPiece, VectorDrawing } from './drawings';

const EMU = 12700;
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
/** The control point distance of a quarter ellipse drawn as a cubic curve. */
const KAPPA = 0.5523;
const MAX_LINE_WIDTH = 20_116_800;

const emu = (points: number) => Math.round(points * EMU);
const size = (points: number) => Math.max(1, emu(points));

function colorXml(hex: string, alpha: number | undefined): string {
	const value = hex.replace('#', '').toUpperCase();
	const opacity =
		alpha !== undefined ? `<a:alpha val="${Math.round(Math.max(0, alpha) * 100_000)}"/>` : '';
	return opacity
		? `<a:srgbClr val="${value}">${opacity}</a:srgbClr>`
		: `<a:srgbClr val="${value}"/>`;
}

/** A PDF dash array as DrawingML dashes, in hundred-thousandths of the line width. */
function dashXml(dash: number[] | undefined, width: number): string {
	if (!dash?.length) return '';
	const values = dash.length % 2 ? [...dash, ...dash] : dash;
	let out = '';
	for (let index = 0; index + 1 < values.length && index < 20; index += 2) {
		const d = Math.max(1, Math.round((values[index]! / width) * 100_000));
		const sp = Math.max(1, Math.round((values[index + 1]! / width) * 100_000));
		out += `<a:ds d="${d}" sp="${sp}"/>`;
	}
	return `<a:custDash>${out}</a:custDash>`;
}

function styleXml(piece: DrawingPiece): string {
	const fill = piece.fill
		? `<a:solidFill>${colorXml(piece.fill, piece.fillAlpha)}</a:solidFill>`
		: '<a:noFill/>';
	if (!piece.stroke) return `${fill}<a:ln><a:noFill/></a:ln>`;
	// PDF draws a zero width as the thinnest line the device can show.
	const width = Math.max(0.25, piece.lineWidth);
	return (
		`${fill}<a:ln w="${Math.min(MAX_LINE_WIDTH, emu(width))}"><a:solidFill>${colorXml(piece.stroke, piece.strokeAlpha)}</a:solidFill>` +
		`${dashXml(piece.dash, width)}</a:ln>`
	);
}

/** The piece as path segments (for a shape holding several pieces). */
function segmentsOf(piece: DrawingPiece): PdfPathSegment[] {
	const { x0, y0, x1, y1 } = piece;
	const geometry = piece.geometry;
	switch (geometry.kind) {
		case 'path':
			return geometry.segments;
		case 'line':
			return geometry.flipH
				? [
						{ op: 'M', x: x1, y: y0 },
						{ op: 'L', x: x0, y: y1 },
					]
				: [
						{ op: 'M', x: x0, y: y0 },
						{ op: 'L', x: x1, y: y1 },
					];
		case 'rect':
			return [
				{ op: 'M', x: x0, y: y0 },
				{ op: 'L', x: x1, y: y0 },
				{ op: 'L', x: x1, y: y1 },
				{ op: 'L', x: x0, y: y1 },
				{ op: 'Z' },
			];
		case 'ellipse': {
			const [cx, cy, rx, ry] = [(x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2];
			const [kx, ky] = [rx * KAPPA, ry * KAPPA];
			return [
				{ op: 'M', x: cx, y: y0 },
				{ op: 'C', x1: cx + kx, y1: y0, x2: x1, y2: cy - ky, x: x1, y: cy },
				{ op: 'C', x1: x1, y1: cy + ky, x2: cx + kx, y2: y1, x: cx, y: y1 },
				{ op: 'C', x1: cx - kx, y1: y1, x2: x0, y2: cy + ky, x: x0, y: cy },
				{ op: 'C', x1: x0, y1: cy - ky, x2: cx - kx, y2: y0, x: cx, y: y0 },
				{ op: 'Z' },
			];
		}
	}
}

/** One `a:path` in a box at (`x`, `y`) of `w` by `h` EMU. */
function pathXml(segments: PdfPathSegment[], x: number, y: number, w: number, h: number): string {
	const pt = (px: number, py: number) => `<a:pt x="${emu(px - x)}" y="${emu(py - y)}"/>`;
	let out = '';
	for (const segment of segments) {
		if (segment.op === 'Z') out += '<a:close/>';
		else if (segment.op === 'C')
			out += `<a:cubicBezTo>${pt(segment.x1, segment.y1)}${pt(segment.x2, segment.y2)}${pt(segment.x, segment.y)}</a:cubicBezTo>`;
		else if (segment.op === 'M') out += `<a:moveTo>${pt(segment.x, segment.y)}</a:moveTo>`;
		else out += `<a:lnTo>${pt(segment.x, segment.y)}</a:lnTo>`;
	}
	return `<a:path w="${w}" h="${h}">${out}</a:path>`;
}

interface Box {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** A shape's `wps:spPr`: its box (`off` relative to `origin`), geometry and style. */
function shapeProperties(pieces: DrawingPiece[], box: Box, origin: { x: number; y: number }) {
	const first = pieces[0]!;
	const geometry = first.geometry;
	const preset = pieces.length === 1 && geometry.kind !== 'path';
	const [cx, cy] = preset ? [emu(box.width), emu(box.height)] : [size(box.width), size(box.height)];
	const flip = preset && geometry.kind === 'line' && geometry.flipH ? ' flipH="1"' : '';
	const shape = preset
		? `<a:prstGeom prst="${geometry.kind}"><a:avLst/></a:prstGeom>`
		: '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>' +
			pieces.map((piece) => pathXml(segmentsOf(piece), box.x, box.y, cx, cy)).join('') +
			'</a:pathLst></a:custGeom>';
	return (
		`<wps:spPr><a:xfrm${flip}><a:off x="${emu(box.x - origin.x)}" y="${emu(box.y - origin.y)}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
		`${shape}${styleXml(first)}</wps:spPr>`
	);
}

const boxOf = (piece: DrawingPiece): Box => ({
	x: piece.x0,
	y: piece.y0,
	width: piece.x1 - piece.x0,
	height: piece.y1 - piece.y0,
});

const lookOf = (piece: DrawingPiece) =>
	[
		piece.fill,
		piece.stroke,
		piece.stroke ? piece.lineWidth.toFixed(2) : '',
		piece.fillAlpha,
		piece.strokeAlpha,
		piece.dash?.join(' '),
	].join('|');

/** The run holding `drawing`, its ids taken from `nextId`. */
export function drawingXml(drawing: VectorDrawing, nextId: () => number): string {
	const id = nextId();
	const pieces = drawing.pieces;
	const look = lookOf(pieces[0]!);
	const together = pieces.length === 1 || pieces.every((piece) => lookOf(piece) === look);
	const box: Box = drawing;
	const preset = pieces.length === 1 && pieces[0]!.geometry.kind !== 'path';
	const cx = preset ? emu(box.width) : size(box.width);
	const cy = preset ? emu(box.height) : size(box.height);
	let graphic: string;
	let requires = 'wps';
	let name = `Shape ${id}`;
	if (together)
		graphic =
			`<a:graphicData uri="${WPS_URI}"><wps:wsp><wps:cNvSpPr/>` +
			`${shapeProperties(pieces, box, box)}<wps:bodyPr/></wps:wsp></a:graphicData>`;
	else {
		requires = 'wpg';
		name = `Group ${id}`;
		graphic =
			`<a:graphicData uri="${WPG_URI}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
			`<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm>` +
			'</wpg:grpSpPr>' +
			pieces
				.map((piece) => {
					const shapeId = nextId();
					return (
						`<wps:wsp><wps:cNvPr id="${shapeId}" name="Shape ${shapeId}"/><wps:cNvSpPr/>` +
						`${shapeProperties([piece], boxOf(piece), box)}<wps:bodyPr/></wps:wsp>`
					);
				})
				.join('') +
			'</wpg:wgp></a:graphicData>';
	}
	return (
		`<w:r><mc:AlternateContent xmlns:mc="${NS.mc}"${requires === 'wpg' ? ` xmlns:wpg="${WPG_URI}"` : ''}><mc:Choice Requires="${requires}"><w:drawing>` +
		`<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id + 1}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">` +
		`<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(box.x)}</wp:posOffset></wp:positionH>` +
		`<wp:positionV relativeFrom="page"><wp:posOffset>${emu(box.y)}</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
		`<wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr/><a:graphic>${graphic}</a:graphic></wp:anchor>` +
		'</w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>'
	);
}
