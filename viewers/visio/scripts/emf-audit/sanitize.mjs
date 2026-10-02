import { LIMITS } from './preflight.mjs';

const ARITY = { m: 2, l: 2, h: 1, v: 1, c: 6, s: 4, q: 4, t: 2, a: 7, z: 0 };
const NUMBER = /[+-]?(?:\d*\.\d+|\d+\.?\d*)(?:[eE][+-]?\d+)?/y;
const numericPaint = new Set([
	'stroke-width',
	'stroke-miterlimit',
	'opacity',
	'fill-opacity',
	'stroke-opacity',
]);
const enums = {
	'fill-rule': ['nonzero', 'evenodd'],
	'stroke-linecap': ['butt', 'round', 'square'],
	'stroke-linejoin': ['miter', 'round', 'bevel'],
};

// Fail closed, rather than dropping nodes and returning a misleading partial drawing.
// This small neutral model contains no element names, CSS, URLs, fonts or raw markup.
export function sanitizeVectorTree(root, limits = LIMITS) {
	let nodes = 0;
	let characters = 0;
	let coordinates = 0;
	const seen = new WeakSet();
	const fail = (detail) => {
		throw new Error(`emf.unsafe-output: ${detail}`);
	};
	const finite = (value, maximum = limits.coordinate * 16) => {
		if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > maximum)
			fail('non-finite or excessive number');
		return value;
	};
	const charge = (value) => {
		if (typeof value === 'string') characters += value.length;
		if (characters > limits.outputChars) fail('output character budget');
	};
	function numbers(value, count) {
		if (typeof value !== 'string') fail('expected number list');
		const parts = value
			.trim()
			.split(/[\s,]+/)
			.map(Number);
		if (parts.length !== count) fail('wrong number-list arity');
		return parts.map((n) => finite(n));
	}
	function geometry(d) {
		if (typeof d !== 'string') fail('missing path');
		let p = 0;
		let command;
		const out = [];
		const skip = () => {
			while (p < d.length && /[\s,]/.test(d[p])) p++;
		};
		while (p < d.length) {
			skip();
			if (p === d.length) break;
			if (/[a-z]/i.test(d[p])) {
				command = d[p++];
				if (!(command.toLowerCase() in ARITY)) fail('unknown path command');
			} else if (!command || command.toLowerCase() === 'z') fail('missing path command');
			if (!out.length && command.toLowerCase() !== 'm') fail('path must start with moveto');
			const values = [];
			for (let i = 0; i < ARITY[command.toLowerCase()]; i++) {
				skip();
				NUMBER.lastIndex = p;
				const match = NUMBER.exec(d);
				if (!match) fail('invalid path operand');
				p = NUMBER.lastIndex;
				values.push(finite(Number(match[0])));
				if (++coordinates > 100_000) fail('coordinate budget');
			}
			if (
				command.toLowerCase() === 'a' &&
				(!values.slice(3, 5).every((n) => n === 0 || n === 1) || values[0] < 0 || values[1] < 0)
			)
				fail('invalid arc');
			out.push({ command, values });
			if (out.length > limits.outputNodes) fail('path command budget');
			if (command === 'M') command = 'L';
			else if (command === 'm') command = 'l';
		}
		return out;
	}
	function visitChildren(children, depth) {
		if (!Array.isArray(children)) fail('children must be an array');
		// Reject before allocating: Array.map preserves a sparse array's large length.
		if (children.length > limits.outputNodes - nodes) fail('child array node budget');
		const items = [];
		for (let i = 0; i < children.length; i++) {
			if (!Object.hasOwn(children, i)) fail('sparse child array');
			items.push(visit(children[i], depth + 1));
		}
		return items;
	}
	function visit(node, depth) {
		if (
			!node ||
			typeof node !== 'object' ||
			seen.has(node) ||
			depth > limits.depth ||
			++nodes > limits.outputNodes
		)
			fail('node/depth/cycle budget');
		seen.add(node);
		if (Object.keys(node).some((key) => !['tag', 'attrs', 'children'].includes(key)))
			fail('unexpected node field or text');
		if (!node.attrs || typeof node.attrs !== 'object' || Array.isArray(node.attrs))
			fail('attributes missing');
		const attrs = node.attrs;
		if (Object.keys(attrs).length > 20) fail('too many attributes');
		Object.values(attrs).forEach(charge);
		if (depth === 0) {
			if (node.tag !== 'svg') fail('root must be svg');
			if (
				Object.keys(attrs).some(
					(key) => !['xmlns', 'width', 'height', 'viewBox', 'style'].includes(key),
				)
			)
				fail('unexpected root attribute');
			if (
				attrs.xmlns !== 'http://www.w3.org/2000/svg' ||
				(attrs.style !== undefined && attrs.style !== 'isolation:isolate')
			)
				fail('root namespace/style');
			const width = finite(attrs.width, limits.outputDimension);
			const height = finite(attrs.height, limits.outputDimension);
			if (width < 1 || height < 1) fail('empty dimensions');
			const box = numbers(attrs.viewBox, 4);
			if (box[0] !== 0 || box[1] !== 0 || box[2] !== width || box[3] !== height)
				fail('viewBox differs from dimensions');
			return {
				kind: 'vector',
				width,
				height,
				items: visitChildren(node.children, depth),
			};
		}
		if (node.tag === 'g') {
			if (Object.keys(attrs).some((key) => key !== 'transform'))
				fail('unsupported group attribute');
			const match =
				typeof attrs.transform === 'string' ? /^matrix\(([^()]*)\)$/.exec(attrs.transform) : null;
			if (attrs.transform !== undefined && !match) fail('unsupported group transform');
			return {
				kind: 'group',
				matrix: match ? numbers(match[1], 6) : [1, 0, 0, 1, 0, 0],
				items: visitChildren(node.children, depth),
			};
		}
		if (
			node.tag !== 'path' ||
			(node.children !== undefined && (!Array.isArray(node.children) || node.children.length))
		)
			fail('only path/group nodes are admitted');
		const paint = {};
		for (const [key, value] of Object.entries(attrs)) {
			if (key === 'd') continue;
			if (['fill', 'stroke'].includes(key)) {
				if (typeof value !== 'string' || !/^(?:none|#[0-9a-f]{6})$/i.test(value))
					fail('nonliteral paint');
				paint[key] = value;
			} else if (numericPaint.has(key)) {
				if (typeof value !== 'string' && typeof value !== 'number') fail('invalid numeric paint');
				const number = finite(Number(value), key.includes('opacity') ? 1 : 1024);
				if (number < 0) fail('negative paint');
				paint[key] = number;
			} else if (enums[key]?.includes(value)) paint[key] = value;
			else fail(`unsupported attribute ${key}`);
		}
		return { kind: 'path', commands: geometry(attrs.d), paint };
	}
	return visit(root, 0);
}
