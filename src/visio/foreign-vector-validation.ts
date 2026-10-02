import {
	VisioForeignVectorError,
	limit,
	unsafe,
	type VisioForeignVector,
	type VisioForeignVectorClip,
	type VisioForeignVectorClipPath,
	type VisioForeignVectorCommand,
	type VisioForeignVectorLimits,
	type VisioForeignVectorMatrix,
	type VisioForeignVectorNode,
	type VisioForeignVectorPaint,
} from './foreign-vector-types.js';
import { checkVectorGraph } from './foreign-vector-graph.js';
import { finite, limitsFor, paint, record } from './foreign-vector-values.js';

const ARITY: Readonly<Record<string, number>> = { M: 2, L: 2, C: 6, Q: 4, A: 7, Z: 0 };
const PAINT = {
	fill: 'fill',
	stroke: 'stroke',
	fillRule: 'fill-rule',
	strokeWidth: 'stroke-width',
	strokeMiterlimit: 'stroke-miterlimit',
	strokeLinecap: 'stroke-linecap',
	strokeLinejoin: 'stroke-linejoin',
	opacity: 'opacity',
	fillOpacity: 'fill-opacity',
	strokeOpacity: 'stroke-opacity',
};
/** Validate untrusted worker/host scenes, returning an independent deeply frozen canonical copy. */
export function validateVisioForeignVector(
	value: unknown,
	options: Partial<VisioForeignVectorLimits> = {},
): VisioForeignVector {
	try {
		return validate(value, limitsFor(options));
	} catch (error) {
		if (error instanceof VisioForeignVectorError) throw error;
		unsafe('Unable to inspect vector scene safely.');
	}
}
function validate(value: unknown, limits: VisioForeignVectorLimits): VisioForeignVector {
	let nodes = 0,
		operands = 0,
		commands = 0,
		characters = 0,
		clipCount = 0;
	const seen = new WeakSet<object>();
	function own(input: unknown, fields: readonly string[], depth: number): Record<string, unknown> {
		if (depth > limits.maxDepth) limit('Scene depth limit exceeded.');
		if (++nodes > limits.maxNodes) limit('Scene node limit exceeded.');
		if (!input || typeof input !== 'object' || seen.has(input))
			unsafe('Repeated or cyclic scene object.');
		seen.add(input);
		return data(input, fields);
	}
	function data(input: unknown, fields: readonly string[]): Record<string, unknown> {
		const output = record(input, fields);
		for (const value of Object.values(output)) {
			characters += typeof value === 'string' ? value.length : 0;
			if (characters > limits.maxCharacters) limit('Scene character limit exceeded.');
		}
		return output;
	}
	function array(input: unknown, maximum: number): unknown[] {
		if (!Array.isArray(input) || Object.getPrototypeOf(input) !== Array.prototype)
			unsafe('Expected plain scene array.');
		if (input.length > maximum) limit('Scene array limit exceeded.');
		if (Reflect.ownKeys(input).length !== input.length + 1)
			unsafe('Sparse or extended scene array.');
		const output: unknown[] = [];
		for (let i = 0; i < input.length; i++) {
			const item = Object.getOwnPropertyDescriptor(input, i);
			if (!item || !('value' in item) || !item.enumerable)
				unsafe('Sparse or accessor scene array.');
			output.push(item.value as unknown);
		}
		return output;
	}
	function base(input: Record<string, unknown>): {
		matrix: VisioForeignVectorMatrix;
		clipIndex?: number;
	} {
		const values = array(input.matrix, 6);
		if (values.length !== 6) unsafe('Invalid matrix arity.');
		const matrix = Object.freeze(
			values.map((item) => finite(item, limits.maxCoordinate)),
		) as unknown as VisioForeignVectorMatrix;
		if (input.clipIndex === undefined) return { matrix };
		if (
			typeof input.clipIndex !== 'number' ||
			!Number.isSafeInteger(input.clipIndex) ||
			input.clipIndex < 0 ||
			input.clipIndex >= clipCount
		)
			unsafe('Invalid clip index.');
		return { matrix, clipIndex: input.clipIndex };
	}
	function path(input: unknown): readonly VisioForeignVectorCommand[] {
		const values = array(input, limits.maxPathCommands - commands);
		const result: VisioForeignVectorCommand[] = [];
		for (const value of values) {
			const command = data(value, ['command', 'values']);
			if (typeof command.command !== 'string' || !Object.hasOwn(ARITY, command.command))
				unsafe('Invalid normalized command.');
			if (!result.length && command.command !== 'M') unsafe('Path must begin with moveto.');
			const count = ARITY[command.command]!;
			const numbers = array(command.values, count);
			if (numbers.length !== count) unsafe('Invalid normalized command arity.');
			if ((operands += count) > limits.maxPathOperands) limit('Scene operand limit exceeded.');
			commands++;
			const copied = numbers.map((item) => finite(item, limits.maxCoordinate));
			if (
				command.command === 'A' &&
				(copied[0]! < 0 ||
					copied[1]! < 0 ||
					![0, 1].includes(copied[3]!) ||
					![0, 1].includes(copied[4]!))
			)
				unsafe('Invalid normalized arc.');
			result.push(
				Object.freeze({
					command: command.command as VisioForeignVectorCommand['command'],
					values: Object.freeze(copied),
				}),
			);
		}
		return Object.freeze(result);
	}
	function literalPaint(input: unknown): VisioForeignVectorPaint {
		const fields = Object.keys(PAINT) as (keyof typeof PAINT)[];
		const value = data(input, fields);
		if (Object.keys(value).length !== fields.length) unsafe('Missing normalized paint field.');
		const attrs: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
		for (const field of fields) {
			if (value[field] === undefined) unsafe('Missing normalized paint value.');
			if (
				['strokeWidth', 'strokeMiterlimit', 'opacity', 'fillOpacity', 'strokeOpacity'].includes(
					field,
				) &&
				typeof value[field] !== 'number'
			)
				unsafe('Normalized paint must be numeric.');
			attrs[PAINT[field]] = value[field];
		}
		return paint(attrs);
	}
	function item(input: unknown, depth: number): VisioForeignVectorNode {
		const node = own(input, ['kind', 'matrix', 'clipIndex', 'items', 'commands', 'paint'], depth);
		if (node.kind === 'path') {
			if (Object.hasOwn(node, 'items')) unsafe('Path contains group fields.');
			return Object.freeze({
				kind: 'path',
				...base(node),
				commands: path(node.commands),
				paint: literalPaint(node.paint),
			});
		}
		if (node.kind !== 'group' || Object.hasOwn(node, 'commands') || Object.hasOwn(node, 'paint'))
			unsafe('Invalid normalized group.');
		return Object.freeze({
			kind: 'group',
			...base(node),
			items: Object.freeze(
				array(node.items, limits.maxNodes - nodes).map((child) => item(child, depth + 1)),
			),
		});
	}
	function clip(input: unknown, depth: number): VisioForeignVectorClip {
		const node = own(input, ['matrix', 'clipIndex', 'items'], depth);
		const items: VisioForeignVectorClipPath[] = [];
		for (const value of array(node.items, limits.maxNodes - nodes)) {
			const child = own(value, ['matrix', 'clipIndex', 'commands', 'clipRule'], depth + 1);
			if (child.clipRule !== 'nonzero' && child.clipRule !== 'evenodd')
				unsafe('Invalid normalized clip rule.');
			items.push(
				Object.freeze({ ...base(child), commands: path(child.commands), clipRule: child.clipRule }),
			);
		}
		return Object.freeze({ ...base(node), items: Object.freeze(items) });
	}
	const root = own(value, ['kind', 'width', 'height', 'items', 'clips'], 0);
	if (root.kind !== 'vector') unsafe('Invalid normalized vector.');
	const width = finite(root.width, limits.maxDimension),
		height = finite(root.height, limits.maxDimension);
	if (width < 1 || height < 1) unsafe('Empty normalized dimensions.');
	const definitions = array(root.clips, limits.maxNodes - nodes);
	clipCount = definitions.length;
	const clips = Object.freeze(definitions.map((value) => clip(value, 1)));
	const items = Object.freeze(
		array(root.items, limits.maxNodes - nodes).map((value) => item(value, 1)),
	);
	const scene: VisioForeignVector = { kind: 'vector', width, height, clips, items };
	checkVectorGraph(scene, limits);
	return Object.freeze(scene);
}
