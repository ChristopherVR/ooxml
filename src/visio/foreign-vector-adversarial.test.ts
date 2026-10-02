import { describe, expect, it, vi } from 'vitest';
import {
	sanitizeVisioForeignVectorTree as sanitize,
	validateVisioForeignVector as validate,
	VisioForeignVectorError,
} from './foreign-vector.js';
import { root, path, group, clip, defs } from './foreign-vector-test-fixtures.js';

const rejected = (input: unknown, code: 'unsafe-vector' | 'vector-limit' = 'unsafe-vector') => {
	expect(() => sanitize(input)).toThrowError(
		expect.objectContaining({ name: 'VisioForeignVectorError', code }),
	);
};
const invalidPaths = [
	'L0 0',
	'M',
	'M0',
	'M,0 0',
	'M0 0,',
	'M0 0,,1 1',
	'M0 0,L1 1',
	'M0 0L1',
	'M0 0L1 1 2',
	'M0 0Z1 2',
	'M0 0B1 1',
	'MNaN 0',
	'MInfinity 0',
	'M0x10 0',
	'M1e999 0',
	'M1e 0',
	'M0 0A-1 1 0 0 0 1 1',
	'M0 0A1 1 0 2 0 1 1',
	'M0 0A1 1 0 +1 0 1 1',
	'M0 0A1 1 0 1.0 0 1 1',
	'M0 0A1 1 0 0 0',
	'M0 0<script/>',
	'M0\u00a00',
];

describe('fail-closed converter tree boundary', () => {
	it.each([
		'script',
		'text',
		'image',
		'use',
		'foreignObject',
		'a',
		'style',
		'filter',
		'animate',
		'svg',
		'rect',
	])('rejects unsupported %s elements without partial rendering', (tag) =>
		rejected(root([path(), { tag, attrs: {}, children: [] }])),
	);
	it.each([
		'onclick',
		'onload',
		'style',
		'class',
		'font-family',
		'filter',
		'mask',
		'href',
		'xlink:href',
		'xml:base',
		'marker-end',
		'stroke-dasharray',
		'id',
		'__proto__',
		'constructor',
	])('rejects attribute %s', (key) => rejected(root([path('M0 0', { [key]: 'payload' })])));
	it.each([
		'url(#a)',
		'url(https://example.com/a)',
		'var(--color)',
		'currentColor',
		'red',
		'#fff',
		'#00000000',
		'inherit',
		'none;stroke:red',
	])('rejects nonliteral paint %s', (fill) => rejected(root([path('M0 0', { fill })])));
	it.each(invalidPaths)('rejects malformed path %s', (d) => rejected(root([path(d)])));
	it.each([null, true, {}, [], '', ' ', '0x10', 'Infinity', 'NaN', '1px', '-1'])(
		'rejects invalid numeric paint %j',
		(value) => rejected(root([path('M0 0', { 'stroke-width': value })])),
	);
	it.each([null, undefined, NaN, Infinity, '20', 0, -1])('rejects invalid root width %j', (width) =>
		rejected(root([], { width })),
	);
	it('rejects invalid numeric paint ranges', () => {
		for (const attrs of [
			{ opacity: -1 },
			{ opacity: 1.1 },
			{ 'stroke-miterlimit': 0 },
			{ 'fill-rule': 'inherit' },
		])
			expect(() => sanitize(root([path('M0 0', attrs)]))).toThrow(VisioForeignVectorError);
	});
	it('rejects invalid root viewport, namespace, style and unknown root fields', () => {
		for (const attrs of [
			{ viewBox: '0 0 21 20' },
			{ viewBox: '1 0 20 20' },
			{ viewBox: '0 0 20 20 1' },
			{ xmlns: 'https://www.w3.org/2000/svg' },
			{ style: 'isolation:isolate;color:red' },
			{ preserveAspectRatio: 'none' },
		])
			rejected(root([], attrs));
		rejected({ ...root(), text: 'hidden' });
		rejected({ ...root(), [Symbol('hidden')]: true });
		rejected(root([{ ...path(), children: [path()] }]));
	});
	it.each([
		'translate(1 2)',
		'matrix(1 0 0 1 0)',
		'matrix(1 0 0 1 0 0 0)',
		'matrix(1 0 0 1 NaN 0)',
		'matrix(1 0 0 1 0 0);color:red',
	])('rejects unsupported transform %s', (transform) =>
		rejected(root([group([path()], { transform })])),
	);
	it('rejects huge and sparse child arrays before visiting or mapping entries', () => {
		const getter = vi.fn(() => path()),
			huge = new Array(0xffffffff);
		Object.defineProperty(huge, '0', { get: getter });
		for (const input of [
			root(huge),
			root([group(huge)]),
			root([defs(huge)]),
			root([defs([clip('c', huge)])]),
		])
			rejected(input, 'vector-limit');
		expect(getter).not.toHaveBeenCalled();
		rejected(root(new Array(5)));
		rejected(root(Object.assign([path()], { unwanted: true })));
		const sparse = [path(), , path()];
		rejected(root(sparse));
	});
	it('rejects accessors without invoking them and never string-coerces values', () => {
		const getter = vi.fn(() => 'M0 0'),
			coercion = vi.fn(() => '1');
		const child = path(),
			input = root([child]);
		Object.defineProperty(child.attrs, 'd', { enumerable: true, get: getter });
		rejected(input);
		rejected(root([path('M0 0', { 'stroke-width': { toString: coercion } })]));
		expect(getter).not.toHaveBeenCalled();
		expect(coercion).not.toHaveBeenCalled();
	});
	it('rejects exotic prototypes, type mismatches and hidden fields', () => {
		for (const input of [
			new Date(),
			[],
			Object.create(root()),
			{ ...root(), attrs: [] },
			{ ...root(), children: {} },
		])
			rejected(input);
		const customArray = [path()];
		Object.setPrototypeOf(customArray, {});
		rejected(root(customArray));
		const hidden = root();
		Object.defineProperty(hidden, 'secret', { value: true });
		rejected(hidden);
	});
	it('rejects cycles and repeated node or child-array references', () => {
		const repeated = path();
		rejected(root([repeated, repeated]));
		const cycle = group([]);
		cycle.children.push(cycle);
		rejected(root([cycle]));
		const same: unknown[] = [];
		rejected(root([group(same), group(same)]));
	});
	it('never includes untrusted exception or source content in errors', () => {
		const input = new Proxy(
			{},
			{
				getPrototypeOf() {
					throw new Error('secret source');
				},
			},
		);
		expect(() => sanitize(input)).toThrow('Unable to inspect converter output safely.');
		expect(() => sanitize(input)).not.toThrow('secret source');
	});
	it('bounds aggregate text, path operands, normalized commands, nodes and depth', () => {
		const cases = [
			{ maxCharacters: 10 },
			{ maxPathOperands: 3 },
			{ maxPathCommands: 1 },
			{ maxNodes: 1 },
		];
		for (const option of cases)
			expect(() => sanitize(root(), option)).toThrowError(
				expect.objectContaining({ code: 'vector-limit' }),
			);
		expect(() => sanitize(root([path('M0 0H1')]), { maxPathOperands: 3 })).toThrowError(
			expect.objectContaining({ code: 'vector-limit' }),
		);
		expect(() => sanitize(root([group([group([path()])])]), { maxDepth: 2 })).toThrowError(
			expect.objectContaining({ code: 'vector-limit' }),
		);
		for (const bad of [0, -1, 1.5, Infinity, NaN, '1', null])
			expect(() => sanitize(root(), { maxNodes: bad as number })).toThrow(VisioForeignVectorError);
		expect(() => sanitize(root(), { unknown: 1 } as object)).toThrow(VisioForeignVectorError);
	});
	it('bounds relative accumulation, reflected controls, arc radius correction and composed transforms', () => {
		for (const d of [
			'M900000 0l200000 0',
			'M0 0Q-900000 0 100000 0T0 0',
			'M0 0A1e-300 100 0 0 0 10 10',
		])
			expect(() => sanitize(root([path(d)]))).toThrow(VisioForeignVectorError);
		const transform = { transform: 'matrix(1000 0 0 1000 0 0)' };
		rejected(
			root([group([group([group([path()], transform)], transform)], transform)]),
			'vector-limit',
		);
		rejected(
			root([group([path('M1000 0L2000 0')], { transform: 'matrix(1000 0 0 1000 0 0)' })]),
			'vector-limit',
		);
	});
});

describe('closed clip graph validation', () => {
	it('rejects duplicate, dangling, non-clip and malformed identities', () => {
		rejected(root([defs([clip('c'), clip('c')])]));
		for (const ref of [
			'url(#missing)',
			'url(https://example.com/#c)',
			'url(data:image/svg+xml,a)',
			'url(#c) url(#d)',
			'#c',
			'url("#c")',
			'none',
		])
			rejected(root([group([path()], { 'clip-path': ref })]));
		for (const id of ['x y', '', '#x', 'x'.repeat(129)]) rejected(root([defs([clip(id)])]));
		rejected(root([defs([path()])]));
		rejected(root([defs([clip('c', [group([path()])])])]));
		rejected(root([defs([clip('c', [path()], { clipPathUnits: 'objectBoundingBox' })])]));
	});
	it('rejects unused self-cycles, mutual cycles and child-reference cycles', () => {
		rejected(root([defs([clip('a', [], { 'clip-path': 'url(#a)' })])]));
		rejected(
			root([
				defs([
					clip('a', [], { 'clip-path': 'url(#b)' }),
					clip('b', [], { 'clip-path': 'url(#a)' }),
				]),
			]),
		);
		rejected(root([defs([clip('a', [path('M0 0', { 'clip-path': 'url(#a)' })])])]));
	});
	it('bounds clip graph depth and repeated resource expansion', () => {
		const chain = Array.from({ length: 10 }, (_, i) =>
			clip(`c${i}`, [path()], i ? { 'clip-path': `url(#c${i - 1})` } : {}),
		);
		expect(() => sanitize(root([defs(chain)]), { maxDepth: 5 })).toThrowError(
			expect.objectContaining({ code: 'vector-limit' }),
		);
		const branch = Array.from({ length: 8 }, (_, i) =>
			clip(
				`c${i}`,
				i
					? [
							path('M0 0', { 'clip-path': `url(#c${i - 1})` }),
							path('M0 0', { 'clip-path': `url(#c${i - 1})` }),
						]
					: [path()],
			),
		);
		expect(() => sanitize(root([defs(branch)]), { maxExpandedNodes: 100 })).toThrowError(
			expect.objectContaining({ code: 'vector-limit' }),
		);
		expect(() =>
			sanitize(root([defs([clip('c')]), group([path()], { 'clip-path': 'url(#c)' })]), {
				maxExpandedOperands: 10,
			}),
		).toThrowError(expect.objectContaining({ code: 'vector-limit' }));
	});
	it('checks clip transforms and geometry at every use site', () => {
		rejected(
			root([
				defs([clip('c', [path('M1000 0L2000 0')])]),
				group([path('M0 0')], { transform: 'matrix(1000 0 0 1000 0 0)', 'clip-path': 'url(#c)' }),
			]),
			'vector-limit',
		);
	});
});
