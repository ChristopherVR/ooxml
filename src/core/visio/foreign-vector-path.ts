import {
	limit,
	unsafe,
	type VisioForeignVectorCommand,
	type VisioForeignVectorLimits,
	type VisioForeignVectorMatrix,
} from './foreign-vector-types.js';
import { finite } from './foreign-vector-values.js';

const ARITY: Readonly<Record<string, number>> = {
	M: 2,
	L: 2,
	H: 1,
	V: 1,
	C: 6,
	S: 4,
	Q: 4,
	T: 2,
	A: 7,
	Z: 0,
};
const NUMBER = /[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y;
const WSP = /[\x20\t\r\n]/;
export interface PathBudget {
	operands: number;
	commands: number;
}
/** Deliberately strict SVG grammar: malformed paths reject the whole drawing, never partial playback.
 * https://www.w3.org/TR/SVG11/paths.html#PathDataBNF
 */
export function pathCommands(
	value: unknown,
	limits: VisioForeignVectorLimits,
	budget: PathBudget,
): readonly VisioForeignVectorCommand[] {
	if (typeof value !== 'string') unsafe('Missing path data.');
	const data = value;
	let position = 0,
		current = '',
		x = 0,
		y = 0,
		startX = 0,
		startY = 0;
	let controlX = 0,
		controlY = 0,
		previous = '';
	const result: VisioForeignVectorCommand[] = [];
	const whitespace = () => {
		while (position < data.length && WSP.test(data[position]!)) position++;
	};
	function operand(separator: boolean, flag: boolean): number {
		whitespace();
		if (separator && data[position] === ',') {
			position++;
			whitespace();
		}
		if (flag) {
			const char = data[position++];
			if (char !== '0' && char !== '1') unsafe('Invalid arc flag.');
			return Number(char);
		}
		NUMBER.lastIndex = position;
		const match = NUMBER.exec(data);
		if (!match) unsafe('Invalid path operand.');
		position = NUMBER.lastIndex;
		return finite(Number(match[0]), limits.maxCoordinate);
	}
	while (true) {
		whitespace();
		if (position === data.length) break;
		const explicit = /^[a-zA-Z]$/.test(data[position]!);
		if (explicit) current = data[position++]!;
		if (!current || !Object.hasOwn(ARITY, current.toUpperCase())) unsafe('Unknown path command.');
		const command = current.toUpperCase(),
			relative = current !== command;
		if (!result.length && command !== 'M') unsafe('Path must begin with moveto.');
		const arity = ARITY[command]!;
		if (budget.operands + arity > limits.maxPathOperands) limit('Path operand limit exceeded.');
		if (budget.commands >= limits.maxPathCommands) limit('Path command limit exceeded.');
		const values: number[] = [];
		for (let i = 0; i < arity; i++)
			values.push(operand(i > 0 || !explicit, command === 'A' && (i === 3 || i === 4)));
		const px = (i: number) => values[i]! + (relative ? x : 0);
		const py = (i: number) => values[i]! + (relative ? y : 0);
		let output: VisioForeignVectorCommand;
		switch (command) {
			case 'M':
			case 'L':
				output = { command, values: [px(0), py(1)] };
				break;
			case 'H':
				output = { command: 'L', values: [px(0), y] };
				break;
			case 'V':
				output = { command: 'L', values: [x, py(0)] };
				break;
			case 'C':
				output = { command: 'C', values: [px(0), py(1), px(2), py(3), px(4), py(5)] };
				break;
			case 'S':
				output = {
					command: 'C',
					values: [
						previous === 'C' || previous === 'S' ? 2 * x - controlX : x,
						previous === 'C' || previous === 'S' ? 2 * y - controlY : y,
						px(0),
						py(1),
						px(2),
						py(3),
					],
				};
				break;
			case 'Q':
				output = { command: 'Q', values: [px(0), py(1), px(2), py(3)] };
				break;
			case 'T':
				output = {
					command: 'Q',
					values: [
						previous === 'Q' || previous === 'T' ? 2 * x - controlX : x,
						previous === 'Q' || previous === 'T' ? 2 * y - controlY : y,
						px(0),
						py(1),
					],
				};
				break;
			case 'A':
				if (values[0]! < 0 || values[1]! < 0) unsafe('Negative arc radius.');
				output = { command: 'A', values: [...values.slice(0, 5), px(5), py(6)] };
				break;
			case 'Z':
				output = { command: 'Z', values: [] };
				break;
			default:
				unsafe('Unknown path command.');
		}
		budget.operands += Math.max(arity, output.values.length);
		budget.commands++;
		if (budget.operands > limits.maxPathOperands) limit('Normalized path operand limit exceeded.');
		for (const item of output.values) finite(item, limits.maxCoordinate);
		if (output.command === 'Z') {
			x = startX;
			y = startY;
			current = '';
		} else {
			x = output.values[output.values.length - 2]!;
			y = output.values[output.values.length - 1]!;
			if (command === 'M') {
				startX = x;
				startY = y;
				current = relative ? 'l' : 'L';
			}
			if (output.command === 'C' || output.command === 'Q') {
				controlX = output.values[output.values.length - 4]!;
				controlY = output.values[output.values.length - 3]!;
			}
		}
		previous = command;
		Object.freeze(output.values);
		result.push(Object.freeze(output));
	}
	return Object.freeze(result);
}

/** Conservative transformed bounds, including accumulated relative coordinates and corrected arc radii. */
export function checkPathBounds(
	commands: readonly VisioForeignVectorCommand[],
	matrix: VisioForeignVectorMatrix,
	max: number,
): void {
	const point = (x: number, y: number) => {
		finite(x, max);
		finite(y, max);
		finite(matrix[0] * x + matrix[2] * y + matrix[4], max);
		finite(matrix[1] * x + matrix[3] * y + matrix[5], max);
	};
	let x = 0,
		y = 0,
		startX = 0,
		startY = 0;
	for (const { command, values: v } of commands) {
		if (command === 'Z') {
			x = startX;
			y = startY;
			continue;
		}
		if (command === 'A') {
			const endX = v[5]!,
				endY = v[6]!;
			point(endX, endY);
			if (v[0] !== 0 && v[1] !== 0 && (x !== endX || y !== endY)) {
				// SVG endpoint-to-center conversion corrects radii that are too small for the endpoints.
				const angle = ((v[2]! % 360) * Math.PI) / 180,
					c = Math.cos(angle),
					s = Math.sin(angle);
				const dx = (x - endX) / 2,
					dy = (y - endY) / 2;
				const scale = Math.max(
					1,
					Math.hypot((c * dx + s * dy) / v[0]!, (-s * dx + c * dy) / v[1]!),
				);
				const rx = finite(v[0]! * scale, max),
					ry = finite(v[1]! * scale, max);
				const bx = 2 * (Math.abs(c) * rx + Math.abs(s) * ry);
				const by = 2 * (Math.abs(s) * rx + Math.abs(c) * ry);
				for (const signX of [-1, 1])
					for (const signY of [-1, 1]) point(x + signX * bx, y + signY * by);
			}
			x = endX;
			y = endY;
		} else {
			for (let i = 0; i < v.length; i += 2) point(v[i]!, v[i + 1]!);
			x = v[v.length - 2]!;
			y = v[v.length - 1]!;
			if (command === 'M') {
				startX = x;
				startY = y;
			}
		}
	}
}
