import { describe, expect, it, vi } from 'vitest';
import {
	sanitizeVisioForeignVectorTree as sanitize,
	validateVisioForeignVector as validate,
	VisioForeignVectorError,
} from './foreign-vector.js';
import { root, path, group, clip, defs } from './foreign-vector-test-fixtures.js';

function canonical(): Record<string, any> {
	return structuredClone(
		sanitize(root([defs([clip('c')]), group([path()], { 'clip-path': 'url(#c)' })])),
	);
}
function rejectMutation(mutate: (scene: Record<string, any>) => void): void {
	const scene = canonical();
	mutate(scene);
	expect(() => validate(scene)).toThrow(VisioForeignVectorError);
}

describe('untrusted canonical foreign-vector boundary', () => {
	it.each(['script', 'attrs', 'xmlns', 'href', 'style', 'rawMarkup', 'image', 'text'])(
		'rejects unexpected %s fields at every structural level',
		(key) => {
			for (const locate of [
				(s: any) => s,
				(s: any) => s.items[0],
				(s: any) => s.items[0].items[0],
				(s: any) => s.clips[0],
				(s: any) => s.clips[0].items[0],
			])
				rejectMutation((s) => {
					locate(s)[key] = 'payload';
				});
		},
	);
	it.each([-1, 1, 1.5, '0', NaN, Infinity, null])(
		'rejects out-of-range or wrongly typed clip index %j',
		(index) =>
			rejectMutation((s) => {
				s.items[0].clipIndex = index;
			}),
	);
	it('rejects clip graph cycles and indexes on clip children', () => {
		rejectMutation((s) => {
			s.clips[0].clipIndex = 0;
		});
		rejectMutation((s) => {
			s.clips[0].items[0].clipIndex = 0;
		});
		rejectMutation((s) => {
			s.clips[0].items[0].clipIndex = 20;
		});
	});
	it.each(['M', 'L', 'C', 'Q', 'A', 'Z'])('rejects wrong arity for %s', (command) =>
		rejectMutation((s) => {
			s.items[0].items[0].commands[0] = { command, values: [1, 2, 3] };
		}),
	);
	it('rejects unnormalized commands, invalid flags, nonfinite values and non-moveto start', () => {
		for (const command of [
			{ command: 'm', values: [0, 0] },
			{ command: 'L', values: [0, 0] },
			{ command: '<script>', values: [] },
			{ command: 'M', values: [Infinity, 0] },
			{ command: 'M', values: ['1', 0] },
		])
			rejectMutation((s) => {
				s.items[0].items[0].commands = [command];
			});
		for (const values of [
			[1, 1, 0, 2, 0, 1, 1],
			[-1, 1, 0, 0, 0, 1, 1],
		])
			rejectMutation((s) => {
				s.items[0].items[0].commands.push({ command: 'A', values });
			});
	});
	it('rejects paint injections, missing values, stringified numbers and invalid clip rules', () => {
		for (const [key, value] of [
			['fill', 'url(#x)'],
			['stroke', 'currentColor'],
			['fillRule', 'inherit'],
			['strokeWidth', '1'],
			['strokeMiterlimit', null],
			['opacity', Infinity],
		])
			rejectMutation((s) => {
				s.items[0].items[0].paint[key as string] = value;
			});
		rejectMutation((s) => {
			delete s.items[0].items[0].paint.stroke;
		});
		rejectMutation((s) => {
			s.clips[0].items[0].clipRule = 'inherit';
		});
		rejectMutation((s) => {
			s.clips[0].items[0].paint = {};
		});
	});
	it('rejects mismatched node variants', () => {
		rejectMutation((s) => {
			s.items[0].commands = [];
		});
		rejectMutation((s) => {
			s.items[0].items[0].items = [];
		});
		rejectMutation((s) => {
			s.items[0].kind = 'image';
		});
		rejectMutation((s) => {
			s.kind = 'svg';
		});
	});
	it('bounds sparse arrays before accessing values at all array-bearing fields', () => {
		const huge = new Array(0xffffffff),
			getter = vi.fn();
		Object.defineProperty(huge, 0, { get: getter });
		for (const mutate of [
			(s: any) => {
				s.items = huge;
			},
			(s: any) => {
				s.clips = huge;
			},
			(s: any) => {
				s.items[0].matrix = huge;
			},
			(s: any) => {
				s.clips[0].items = huge;
			},
			(s: any) => {
				s.items[0].items[0].commands = huge;
			},
			(s: any) => {
				s.items[0].items[0].commands[0].values = huge;
			},
		])
			rejectMutation(mutate);
		expect(getter).not.toHaveBeenCalled();
		rejectMutation((s) => {
			s.items = new Array(2);
		});
		rejectMutation((s) => {
			s.items[0].matrix = [1, 0, 0, 1, 0];
		});
	});
	it('rejects repeated/cyclic structural objects and accessors without running getters', () => {
		rejectMutation((s) => {
			s.items.push(s.items[0]);
		});
		rejectMutation((s) => {
			s.items[0].items.push(s.items[0]);
		});
		const getter = vi.fn();
		rejectMutation((s) => {
			Object.defineProperty(s.items[0], 'matrix', { enumerable: true, get: getter });
		});
		expect(getter).not.toHaveBeenCalled();
	});
	it('rejects transform and use-site amplification after transport', () => {
		rejectMutation((s) => {
			s.items[0].matrix = [1_000_000, 0, 0, 1_000_000, 0, 0];
		});
		rejectMutation((s) => {
			s.clips[0].items[0].matrix = [1_000_000, 0, 0, 1_000_000, 0, 0];
		});
	});
	it('enforces aggregate canonical node, text, operand, command and expanded-work limits', () => {
		for (const options of [
			{ maxNodes: 1 },
			{ maxCharacters: 1 },
			{ maxPathOperands: 1 },
			{ maxPathCommands: 1 },
			{ maxDepth: 1 },
			{ maxExpandedNodes: 1 },
			{ maxExpandedOperands: 1 },
			{ maxExpandedCommands: 1 },
		])
			expect(() => validate(canonical(), options)).toThrowError(
				expect.objectContaining({ code: 'vector-limit' }),
			);
	});
	it('rejects generated command mutations deterministically and preserves input after failure', () => {
		for (let seed = 0; seed < 100; seed++) {
			const scene = canonical(),
				before = structuredClone(scene);
			const shape = scene.items[0].items[0];
			shape.commands[0].values[seed % 2] =
				seed % 3 === 0 ? Infinity : seed % 3 === 1 ? 1_000_001 + seed : `url(#${seed})`;
			const mutated = structuredClone(scene);
			expect(() => validate(scene)).toThrow(VisioForeignVectorError);
			expect(scene).toEqual(mutated);
			expect(before.items).not.toEqual(scene.items);
		}
	});
	it('charges zero-operand close commands on every resource use', () => {
		const input = root([
			defs([clip('c', [path('M0 0' + 'Z'.repeat(100))])]),
			group([path('M0 0')], { 'clip-path': 'url(#c)' }),
		]);
		expect(() => sanitize(input, { maxExpandedCommands: 150 })).toThrowError(
			expect.objectContaining({ code: 'vector-limit' }),
		);
		const scene = sanitize(input);
		expect(() => validate(scene, { maxExpandedCommands: 150 })).toThrowError(
			expect.objectContaining({ code: 'vector-limit' }),
		);
	});
	it('counts retained string values consistently without charging generated object keys', () => {
		const many = root(Array.from({ length: 10_000 }, () => path('')));
		const scene = sanitize(many);
		expect(validate(scene)).toEqual(scene);
		expect(() =>
			sanitize(root(Array.from({ length: 100 }, () => path(''))), { maxCharacters: 1000 }),
		).toThrowError(expect.objectContaining({ code: 'vector-limit' }));
	});
	it('rejects long malformed numeric tokens with bounded linear token scans', () => {
		const malformed = '9'.repeat(100_000) + 'x';
		expect(() => sanitize(root([path('M0 0', { 'stroke-width': malformed })]))).toThrow(
			VisioForeignVectorError,
		);
		expect(() => sanitize(root([], { viewBox: malformed }))).toThrow(VisioForeignVectorError);
	});
});
