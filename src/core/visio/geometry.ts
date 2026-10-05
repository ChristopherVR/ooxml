import { splineGeometry } from './spline-geometry.js';
import { geometryRow } from './complex-geometry.js';
import {
	roundedRectanglePath,
	roundedOrthogonalPath,
	type RectanglePoint,
} from './rounded-geometry.js';
import type { VisioGeometry, VisioMatrix } from './model.js';
import { number, type Cells, type Report, type Sheet } from './sheet.js';

const clean = (value: number) => (Math.abs(value) < 1e-12 ? 0 : Number(value.toFixed(9)));
const point = (x: number, y: number) => `${clean(x)} ${clean(y)}`;
/** Visio Pin/LocPin, angle and flips, mapping shape coordinates to its parent's coordinates. */
export function shapeTransform(
	cells: Cells,
	width: number,
	height: number,
	report?: Report,
): VisioMatrix {
	return transform(
		number(cells, 'PinX', width / 2, report),
		number(cells, 'PinY', height / 2, report),
		number(cells, 'LocPinX', width / 2, report),
		number(cells, 'LocPinY', height / 2, report),
		number(cells, 'Angle', 0, report),
		number(cells, 'FlipX', 0, report) !== 0,
		number(cells, 'FlipY', 0, report) !== 0,
	);
}
export function transform(
	x: number,
	y: number,
	localX: number,
	localY: number,
	angle: number,
	flipX = false,
	flipY = false,
): VisioMatrix {
	const cos = Math.cos(angle),
		sin = Math.sin(angle);
	const a = cos * (flipX ? -1 : 1),
		b = sin * (flipX ? -1 : 1);
	const c = -sin * (flipY ? -1 : 1),
		d = cos * (flipY ? -1 : 1);
	return [a, b, c, d, x - a * localX - c * localY, y - b * localX - d * localY];
}
function ellipticalArc(
	startX: number,
	startY: number,
	x: number,
	y: number,
	midX: number,
	midY: number,
	angle: number,
	ratio: number,
): string | undefined {
	if (!(ratio > 0)) return undefined;
	const cos = Math.cos(angle),
		sin = Math.sin(angle);
	const normalize = (px: number, py: number) =>
		[px * cos + py * sin, (-px * sin + py * cos) * ratio] as const;
	const [ax, ay] = normalize(startX, startY),
		[bx, by] = normalize(midX, midY),
		[cx, cy] = normalize(x, y);
	const denominator = 2 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by));
	if (Math.abs(denominator) < 1e-12) return `L ${point(x, y)}`;
	const aa = ax * ax + ay * ay,
		bb = bx * bx + by * by,
		cc = cx * cx + cy * cy;
	const ux = (aa * (by - cy) + bb * (cy - ay) + cc * (ay - by)) / denominator;
	const uy = (aa * (cx - bx) + bb * (ax - cx) + cc * (bx - ax)) / denominator;
	const radius = Math.hypot(ax - ux, ay - uy);
	const tau = Math.PI * 2;
	const start = Math.atan2(ay - uy, ax - ux),
		mid = Math.atan2(by - uy, bx - ux),
		end = Math.atan2(cy - uy, cx - ux);
	const span = (end - start + tau) % tau,
		middle = (mid - start + tau) % tau;
	const sweep = middle <= span ? 1 : 0;
	const large = (sweep ? span : tau - span) > Math.PI ? 1 : 0;
	if (!Number.isFinite(radius) || radius > 1e9) return undefined;
	return `A ${point(radius, radius / ratio)} ${clean((angle * 180) / Math.PI)} ${large} ${sweep} ${point(x, y)}`;
}
/** Converts the supported cached Geometry rows into inert SVG path data. */
export function geometryPaths(
	sheet: Sheet,
	width: number,
	height: number,
	report: Report,
	consume: () => void,
	consumeWork: (units: number) => void = () => {},
): VisioGeometry[] {
	const geometry: VisioGeometry[] = [];
	const rounding = number(sheet.cells, 'Rounding', 0, report);
	if (rounding < 0)
		report('invalid-corner-rounding', 'Negative saved corner rounding was ignored.');
	for (const section of sheet.sections.values()) {
		if (
			section.name !== 'Geometry' ||
			section.deleted ||
			number(section.cells, 'NoShow', 0, report)
		)
			continue;
		const commands: string[] = [];
		const rows = [...section.rows.values()]
			.filter((row) => !row.deleted)
			.sort((a, b) => Number(a.index) - Number(b.index));
		let linePoints: RectanglePoint[] | undefined =
			rows.length >= 3 &&
			rows.every((row, i) =>
				i === 0
					? ['MoveTo', 'RelMoveTo'].includes(row.type)
					: ['LineTo', 'RelLineTo'].includes(row.type),
			)
				? []
				: undefined;
		let currentX = 0,
			currentY = 0,
			currentValid = false,
			supported = true;
		for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
			const row = rows[rowIndex]!;
			consume();
			if (row.type === 'SplineStart') {
				let end = rowIndex + 1;
				while (end < rows.length && rows[end]!.type === 'SplineKnot') {
					consume();
					end++;
				}
				if (end - rowIndex - 1 > 254) {
					report(
						'geometry-limit',
						'A spline exceeds 256 control points; its geometry section was omitted.',
					);
					supported = false;
				} else if (!currentValid || rowIndex === 0 || rows[rowIndex - 1]!.type === 'Ellipse') {
					report(
						'invalid-spline',
						'A spline requires a preceding endpoint row; its geometry section was omitted.',
					);
					supported = false;
				} else {
					const spline = splineGeometry(
						row,
						rows.slice(rowIndex + 1, end),
						currentX,
						currentY,
						report,
						consume,
						consumeWork,
					);
					if (spline) {
						commands.push(spline.path);
						currentX = spline.endX;
						currentY = spline.endY;
						currentValid = true;
					} else {
						supported = false;
						currentValid = false;
					}
				}
				rowIndex = end - 1;
				continue;
			}
			const relative = row.type.startsWith('Rel');
			const xScale = relative ? width : 1,
				yScale = relative ? height : 1;
			const cell = (name: string, fallback = 0) =>
				number(row.cells, name, fallback, (code, message, context) => {
					// A defaulted coordinate must not turn an invalid path into a proven rounded path.
					linePoints = undefined;
					report(code, message, context);
				});
			if (!row.cells.has('X') || !row.cells.has('Y')) linePoints = undefined;
			const x = cell('X') * xScale,
				y = cell('Y') * yScale;
			linePoints?.push([x, y]);
			switch (row.type) {
				case 'MoveTo':
				case 'RelMoveTo':
					commands.push(`M ${point(x, y)}`);
					break;
				case 'LineTo':
				case 'RelLineTo':
					commands.push(`L ${point(x, y)}`);
					break;
				case 'RelQuadBezTo':
					commands.push(`Q ${point(cell('A') * width, cell('B') * height)} ${point(x, y)}`);
					break;
				case 'RelCubBezTo':
					commands.push(
						`C ${point(cell('A') * width, cell('B') * height)} ${point(cell('C') * width, cell('D') * height)} ${point(x, y)}`,
					);
					break;
				case 'ArcTo': {
					const bow = cell('A'),
						chord = Math.hypot(x - currentX, y - currentY);
					if (Math.abs(bow) < 1e-12 || chord < 1e-12) commands.push(`L ${point(x, y)}`);
					else {
						const radius = (chord * chord) / (8 * Math.abs(bow)) + Math.abs(bow) / 2;
						commands.push(
							`A ${point(radius, radius)} 0 ${Math.abs(bow) > chord / 2 ? 1 : 0} ${bow < 0 ? 1 : 0} ${point(x, y)}`,
						);
					}
					break;
				}
				case 'EllipticalArcTo':
				case 'RelEllipticalArcTo': {
					const command = ellipticalArc(
						currentX,
						currentY,
						x,
						y,
						cell('A') * xScale,
						cell('B') * yScale,
						cell('C'),
						cell('D', 1),
					);
					if (command) commands.push(command);
					else {
						supported = false;
						report(
							'invalid-geometry',
							'Invalid elliptical arc parameters; its geometry section was omitted.',
						);
					}
					break;
				}
				case 'Ellipse': {
					const a = cell('A') - x,
						b = cell('B') - y;
					const c = cell('C') - x,
						d = cell('D') - y;
					const rx = Math.hypot(a, b),
						ry = Math.hypot(c, d);
					if (Math.abs(a * c + b * d) > 1e-7 * Math.max(1, rx * ry)) {
						supported = false;
						report(
							'unsupported-ellipse',
							'Non-orthogonal ellipse control axes are unsupported; its geometry section was omitted.',
						);
						break;
					}
					const rotation = clean((Math.atan2(b, a) * 180) / Math.PI);
					commands.push(
						`M ${point(x + a, y + b)} A ${point(rx, ry)} ${rotation} 0 1 ${point(x - a, y - b)} A ${point(rx, ry)} ${rotation} 0 1 ${point(x + a, y + b)} Z`,
					);
					break;
				}
				case 'PolylineTo':
				case 'PolyLineTo':
				case 'NURBSTo': {
					if (row.type === 'NURBSTo' && (!currentValid || rows[rowIndex - 1]?.type === 'Ellipse')) {
						report(
							'invalid-geometry',
							'NURBSTo requires a usable cached preceding endpoint; its geometry section was omitted.',
						);
						supported = false;
						break;
					}
					const suffix = geometryRow(
						row,
						currentX,
						currentY,
						width,
						height,
						report,
						consume,
						consumeWork,
					);
					if (suffix) commands.push(suffix);
					else supported = false;
					break;
				}
				default:
					supported = false;
					report(
						'unsupported-geometry',
						`Geometry row ${row.type || '(missing type)'} is unsupported; its geometry section was omitted.`,
					);
			}
			currentValid =
				Number.isFinite(number(row.cells, 'X', NaN)) &&
				Number.isFinite(number(row.cells, 'Y', NaN)) &&
				Number.isFinite(x) &&
				Number.isFinite(y);
			currentX = x;
			currentY = y;
		}
		if (!supported) continue;
		const rectangle = rounding > 0 ? roundedRectanglePath(linePoints, rounding, point) : undefined;
		const open =
			rounding > 0 && !rectangle ? roundedOrthogonalPath(linePoints, rounding, point) : undefined;
		const rounded = rectangle ?? open?.path;
		if (rounded) {
			// Charge the extra commands beyond the already charged source rows.
			for (let i = 0; i < (rectangle ? 5 : open!.extraCommands); i++) consume();
		} else if (rounding > 0 && commands.length)
			report(
				'unsupported-corner-rounding',
				'Corner rounding requires a closed axis-aligned rectangle or an open orthogonal line chain with space for the saved radius; this geometry was left unchanged.',
			);
		// SVG implicitly closes filled paths; leave strokes open unless the source returns to its origin.
		if (commands.length)
			geometry.push({
				path: rounded ?? commands.join(' '),
				fill: !number(section.cells, 'NoFill', 0, report),
				stroke: !number(section.cells, 'NoLine', 0, report),
			});
	}
	return geometry;
}
