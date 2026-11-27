/**
 * 1-D connectors: the masters they are instances of, and the page shape with its route, its glue
 * formulas and its label. Routes are written out (Visio keeps a saved route until a glued shape
 * moves), so the drawing looks the same in every reader.
 */
import type { MasterDef } from './masters';
import { styleAttrs } from './shape';
import type { StyleBook } from './styles';
import { attrs, cell, row, section } from './xml';

export type Direction = 'n' | 'e' | 's' | 'w';
export interface Point {
	x: number;
	y: number;
}
/** A connection point of a shape on the page. */
export interface Port extends Point {
	shapeId: number;
	/** Zero-based Connection row. */
	index: number;
	dir: Direction;
}

/** Corner radius of a routed connector, in inches. */
const CORNER = 0.05;
const BEGIN = { x: 0.5, y: 1.5 };
const END = { x: 1.5, y: 0.5 };

/**
 * A connector master: a routable 1-D shape (ObjType 2) whose line look comes from its style. The
 * cells are the documented ShapeSheet cells of a connector: guarded 1-D transform, a text block
 * that follows the text, and glue to connection points.
 */
export function connectorMaster(name: string, prompt: string, style: string): MasterDef {
	return {
		name,
		prompt,
		size: { w: 2, h: 2 },
		shape: { style, paths: [] },
		custom: (styles: StyleBook, id: number) => {
			const w = END.x - BEGIN.x;
			const h = END.y - BEGIN.y;
			const cells = [
				cell('PinX', (BEGIN.x + END.x) / 2, 'GUARD((BeginX+EndX)/2)'),
				cell('PinY', (BEGIN.y + END.y) / 2, 'GUARD((BeginY+EndY)/2)'),
				cell('Width', w, 'GUARD(EndX-BeginX)'),
				cell('Height', h, 'GUARD(EndY-BeginY)'),
				cell('LocPinX', w / 2, 'GUARD(Width*0.5)'),
				cell('LocPinY', h / 2, 'GUARD(Height*0.5)'),
				cell('Angle', 0, 'GUARD(0DA)'),
				cell('FlipX', 0, 'GUARD(FALSE)'),
				cell('FlipY', 0, 'GUARD(FALSE)'),
				cell('ResizeMode', 0),
				cell('BeginX', BEGIN.x),
				cell('BeginY', BEGIN.y),
				cell('EndX', END.x),
				cell('EndY', END.y),
				cell('TxtPinX', w / 2, 'Width*0.5'),
				cell('TxtPinY', h / 2, 'Height*0.5'),
				cell('TxtWidth', 0.6, 'MAX(TEXTWIDTH(TheText),5*Char.Size)'),
				cell('TxtHeight', 0.2, 'TEXTHEIGHT(TheText,TxtWidth)'),
				cell('TxtLocPinX', 0.3, 'TxtWidth*0.5'),
				cell('TxtLocPinY', 0.1, 'TxtHeight*0.5'),
				cell('TxtAngle', 0),
				cell('LockCalcWH', 1),
				cell('NoAlignBox', 1),
				cell('DynFeedback', 2),
				cell('GlueType', 2),
				cell('ObjType', 2),
				cell('NoLiveDynamics', 1),
				cell('ShapeSplittable', 1),
				cell('ShapeRouteStyle', 1),
				cell('ConLineRouteExt', 1),
			];
			const geometry = section(
				'Geometry',
				[
					// Relative rows: an instance of any length is right before it writes its own route.
					row({ T: 'RelMoveTo', IX: 1 }, [cell('X', 0), cell('Y', 0)]),
					row({ T: 'RelLineTo', IX: 2 }, [cell('X', 1), cell('Y', 1)]),
				],
				{
					ix: 0,
					cells: [
						cell('NoFill', 1),
						cell('NoLine', 0),
						cell('NoShow', 0),
						cell('NoSnap', 0),
						cell('NoQuickDrag', 0),
					],
				},
			);
			return `<Shape${attrs({ ID: id, Type: 'Shape' })}${styleAttrs(style, styles)}>${cells.join('')}${geometry}</Shape>`;
		},
	};
}

const horizontal = (dir: Direction): boolean => dir === 'e' || dir === 'w';

/** A right-angle route between two ports; `via` gives the corners when the default is not wanted. */
export function route(from: Port, to: Port, via?: readonly Point[]): Point[] {
	const a = { x: from.x, y: from.y };
	const b = { x: to.x, y: to.y };
	if (via) return [a, ...via, b];
	const same = (p: number, q: number) => Math.abs(p - q) < 1e-9;
	if (horizontal(from.dir) && horizontal(to.dir)) {
		if (same(a.y, b.y)) return [a, b];
		const mid = (a.x + b.x) / 2;
		return [a, { x: mid, y: a.y }, { x: mid, y: b.y }, b];
	}
	if (!horizontal(from.dir) && !horizontal(to.dir)) {
		if (same(a.x, b.x)) return [a, b];
		const mid = (a.y + b.y) / 2;
		return [a, { x: a.x, y: mid }, { x: b.x, y: mid }, b];
	}
	if (horizontal(from.dir)) return [a, { x: b.x, y: a.y }, b];
	return [a, { x: a.x, y: b.y }, b];
}

function assertOrthogonal(points: readonly Point[], what: string): void {
	for (let index = 1; index < points.length; index++) {
		const p = points[index - 1];
		const q = points[index];
		if (!p || !q) continue;
		if (Math.abs(p.x - q.x) > 1e-6 && Math.abs(p.y - q.y) > 1e-6)
			throw new Error(`${what}: segment ${index} is not horizontal or vertical.`);
	}
}

/** The point at a distance along a polyline, and whether the segment there is horizontal. */
function pointAt(points: readonly Point[], distance: number): Point & { horizontal: boolean } {
	let left = distance;
	for (let index = 1; index < points.length; index++) {
		const p = points[index - 1];
		const q = points[index];
		if (!p || !q) continue;
		const length = Math.hypot(q.x - p.x, q.y - p.y);
		if (left <= length || index === points.length - 1) {
			const t = length ? Math.min(1, left / length) : 0;
			return {
				x: p.x + (q.x - p.x) * t,
				y: p.y + (q.y - p.y) * t,
				horizontal: Math.abs(q.y - p.y) < 1e-9,
			};
		}
		left -= length;
	}
	return { ...(points[0] ?? { x: 0, y: 0 }), horizontal: true };
}

const pathLength = (points: readonly Point[]): number =>
	points.reduce((total, point, index) => {
		const before = points[index - 1];
		return before ? total + Math.hypot(point.x - before.x, point.y - before.y) : total;
	}, 0);

export interface ConnectorLabel {
	text: string;
	/** Inches from the begin point along the route; the middle of the route when absent. */
	offset?: number;
	/** Label width in inches when the estimate is not good enough. */
	width?: number;
	/** Sit on the line (hiding it behind the label) rather than beside it. */
	onLine?: boolean;
}

export interface ConnectorSpec {
	id: number;
	master: { id: number; name: string };
	name: string;
	from: Port;
	to: Port;
	points: readonly Point[];
	label?: ConnectorLabel;
	/** Extra local cells, such as the layer membership. */
	cells?: readonly string[];
	/** Leave the corners square. */
	square?: boolean;
}

/** The page shape of one glued connector and its two Connect records. */
export function connectorShape(spec: ConnectorSpec): { shape: string; connects: string } {
	const { from, to, points } = spec;
	assertOrthogonal(points, spec.name);
	const w = to.x - from.x;
	const h = to.y - from.y;
	const glue = (port: Port) =>
		`PAR(PNT(Sheet.${port.shapeId}!Connections.X${port.index + 1},Sheet.${port.shapeId}!Connections.Y${port.index + 1}))`;
	const cells = [
		cell('PinX', (from.x + to.x) / 2, 'Inh'),
		cell('PinY', (from.y + to.y) / 2, 'Inh'),
		cell('Width', w, 'GUARD(EndX-BeginX)'),
		cell('Height', h, 'GUARD(EndY-BeginY)'),
		cell('LocPinX', w / 2, 'Inh'),
		cell('LocPinY', h / 2, 'Inh'),
		cell('BeginX', from.x, glue(from)),
		cell('BeginY', from.y, glue(from)),
		cell('EndX', to.x, glue(to)),
		cell('EndY', to.y, glue(to)),
		cell('BegTrigger', 2, `_XFTRIGGER(Sheet.${from.shapeId}!EventXFMod)`),
		cell('EndTrigger', 2, `_XFTRIGGER(Sheet.${to.shapeId}!EventXFMod)`),
	];
	let text = '';
	if (spec.label) {
		const total = pathLength(points);
		const at = pointAt(points, spec.label.offset ?? total / 2);
		// A rough advance of 0.075 inch per character at 8 points keeps the backdrop snug.
		const width = spec.label.width ?? Math.max(0.3, spec.label.text.length * 0.075 + 0.12);
		// Beside the line: above a horizontal run, to the right of a vertical one.
		const beside = spec.label.onLine
			? { x: 0, y: 0 }
			: at.horizontal
				? { x: 0, y: 0.11 }
				: { x: width / 2 + 0.05, y: 0 };
		cells.push(
			cell('TxtPinX', at.x + beside.x - from.x),
			cell('TxtPinY', at.y + beside.y - from.y),
			cell('TxtWidth', width),
			cell('TxtHeight', 0.2),
			cell('TxtLocPinX', width / 2, 'Inh'),
			cell('TxtLocPinY', 0.1, 'Inh'),
		);
		// Only a label that sits on the line needs the style's backdrop to hide it.
		if (!spec.label.onLine) cells.push(cell('TextBkgnd', 0));
		text = `<Text>${spec.label.text}</Text>`;
	}
	// Corners are rounded on the shape, not in the style: a straight line has none to round.
	if (points.length > 2 && !spec.square) cells.push(cell('Rounding', CORNER));
	cells.push(...(spec.cells ?? []));
	const rows = points.map((point, index) =>
		row({ T: index ? 'LineTo' : 'MoveTo', IX: index + 1 }, [
			cell('X', point.x - from.x),
			cell('Y', point.y - from.y),
		]),
	);
	const head = attrs({
		ID: spec.id,
		NameU: spec.name,
		Name: spec.name,
		Type: 'Shape',
		Master: spec.master.id,
	});
	const shape = `<Shape${head}>${cells.join('')}${section('Geometry', rows, { ix: 0 })}${text}</Shape>`;
	const connect = (port: Port, end: 'Begin' | 'End') =>
		`<Connect${attrs({
			FromSheet: spec.id,
			FromCell: `${end}X`,
			FromPart: end === 'Begin' ? 9 : 12,
			ToSheet: port.shapeId,
			ToCell: `Connections.X${port.index + 1}`,
			ToPart: 100 + port.index,
		})}/>`;
	return { shape, connects: connect(from, 'Begin') + connect(to, 'End') };
}
