import { ooxmlArcToSvg } from '../geometry/ooxml-arc';
import type { DrawingPath } from './types';

/** DrawingML arcs share the core preset evaluator's conversion, including full-circle splits. */
export function drawingPathData(path: DrawingPath, width: number, height: number): string {
	const sx = path.width ? width / path.width : 1;
	const sy = path.height ? height / path.height : 1;
	const out: string[] = [];
	let px = 0;
	let py = 0;
	let startX = 0;
	let startY = 0;
	for (const command of path.commands) {
		if (command.op === 'A') {
			const arc = ooxmlArcToSvg(
				command.wR * sx,
				command.hR * sy,
				command.startAngle * 60000,
				command.swingAngle * 60000,
				px,
				py,
			);
			if (arc) {
				out.push(arc.svg);
				px = arc.endX;
				py = arc.endY;
			}
		} else if (command.op === 'Z') {
			out.push('Z');
			px = startX;
			py = startY;
		} else {
			px = command.x * sx;
			py = command.y * sy;
			if (command.op === 'M' || command.op === 'L') {
				out.push(`${command.op}${px} ${py}`);
				if (command.op === 'M') {
					startX = px;
					startY = py;
				}
			} else if (command.op === 'C')
				out.push(
					`C${command.x1 * sx} ${command.y1 * sy} ${command.x2 * sx} ${command.y2 * sy} ${px} ${py}`,
				);
			else if (command.op === 'Q') out.push(`Q${command.x1 * sx} ${command.y1 * sy} ${px} ${py}`);
		}
	}
	return out.join(' ');
}
