import { normalizedNurbs, type NurbsControl } from './nurbs.js';
import { number, type Row, type Report } from './sheet.js';

// Microsoft defines one cached control point per row, with implicit unit weights:
// https://learn.microsoft.com/en-us/office/client-developer/visio/splinestart-row-geometry-section
// https://learn.microsoft.com/en-us/office/client-developer/visio/splineknot-row-geometry-section
// Only contiguous, cached, clamped sequences are supported. Never evaluate formulas.
export interface SplineResult {
	path: string;
	endX: number;
	endY: number;
}
export function splineGeometry(
	start: Row,
	following: readonly Row[],
	startX: number,
	startY: number,
	report: Report,
	consume: () => void,
	consumeWork: (units: number) => void = () => {},
): SplineResult | undefined {
	const invalid = (message: string) => {
		report('invalid-spline', `${message}; its geometry section was omitted.`);
		return undefined;
	};
	if (
		start.type !== 'SplineStart' ||
		following.length === 0 ||
		following.length > 254 ||
		following.some((row) => row.type !== 'SplineKnot')
	)
		return invalid('A spline requires a SplineStart and 1 to 254 contiguous SplineKnot rows');
	const read = (row: Row, name: string) => number(row.cells, name, NaN, report);
	const controls: NurbsControl[] = [
		{ x: startX, y: startY, weight: 1 },
		{ x: read(start, 'X'), y: read(start, 'Y'), weight: 1 },
	];
	const knots = [read(start, 'B'), read(start, 'A')];
	for (const row of following) {
		controls.push({ x: read(row, 'X'), y: read(row, 'Y'), weight: 1 });
		knots.push(read(row, 'A'));
	}
	knots.push(read(start, 'C'));
	const path = normalizedNurbs(
		controls,
		knots,
		read(start, 'D'),
		report,
		consume,
		'SplineStart/SplineKnot',
		consumeWork,
	);
	if (!path) return;
	const end = controls.at(-1)!;
	return { path, endX: end.x, endY: end.y };
}
