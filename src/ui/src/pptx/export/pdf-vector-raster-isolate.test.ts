import { describe, expect, it } from 'vitest';

import {
	isolatedMarkup,
	isolatedVisibilities,
	type IsolatedTemplate,
	type IsolatedVisibilityTree,
} from './pdf-vector-raster-isolate';

/**
 * 0 stage
 * ├─ 1 shape A (the target)
 * │  ├─ 2 text
 * │  │  └─ 3 run
 * │  └─ 4 stylesheet `visibility: hidden`
 * │     └─ 5 inherits hidden
 * └─ 6 shape B, stylesheet `visibility: hidden`
 *    └─ 7 stylesheet `visibility: visible` inside a hidden group
 */
const parents = [-1, 0, 1, 2, 1, 4, 0, 6];
const tree: IsolatedVisibilityTree = {
	parents,
	visibility: ['visible', 'visible', 'visible', 'visible', 'hidden', 'hidden', 'hidden', 'visible'],
	declared: parents.map(() => ''),
};

describe('isolatedVisibilities', () => {
	it('subtree: shows the target and what inherits from it, keeps stylesheet values', () => {
		expect(isolatedVisibilities(tree, 1, 'subtree')).toEqual([
			'hidden',
			'visible',
			'visible',
			'visible',
			'hidden',
			'hidden',
			'hidden',
			// A value of its own differing from its parent's holds under the cascade.
			'visible',
		]);
	});

	it('self and shadow: hide the target children, and what inherits from them', () => {
		for (const mode of ['self', 'shadow'] as const)
			expect(isolatedVisibilities(tree, 1, mode)).toEqual([
				'hidden',
				'visible',
				'hidden',
				'hidden',
				'hidden',
				'hidden',
				'hidden',
				'visible',
			]);
	});

	it('shows a target that is itself hidden, with what inherits from it', () => {
		expect(isolatedVisibilities(tree, 4, 'subtree').slice(4, 6)).toEqual(['visible', 'visible']);
	});

	it('leaves the stage visible when the stage is the target', () => {
		expect(isolatedVisibilities(tree, 0, 'subtree').slice(0, 4)).toEqual([
			'visible',
			'visible',
			'visible',
			'visible',
		]);
		expect(isolatedVisibilities(tree, 0, 'self').slice(0, 2)).toEqual(['visible', 'hidden']);
	});

	it('lets an element with no inlined value inherit', () => {
		const bare: IsolatedVisibilityTree = {
			parents: [-1, 0, 1],
			visibility: ['visible', 'visible', ''],
			declared: ['', '', ''],
		};
		expect(isolatedVisibilities(bare, 1, 'subtree')).toEqual(['hidden', 'visible', 'visible']);
		expect(isolatedVisibilities(bare, 1, 'self')).toEqual(['hidden', 'visible', 'hidden']);
	});

	/**
	 * 0 stage
	 * ├─ 1 target
	 * │  └─ 2 child
	 * │     └─ 3 grandchild, `style="visibility: visible"`
	 * └─ 4 sibling shape, `style="visibility: visible"`
	 *    └─ 5 its text
	 */
	const declaredTree: IsolatedVisibilityTree = {
		parents: [-1, 0, 1, 2, 0, 4],
		visibility: ['visible', 'visible', 'visible', 'visible', 'visible', 'visible'],
		declared: ['', '', '', 'visible', 'visible', ''],
	};

	it('keeps a value the element declares inline, as the capture of the live stage does', () => {
		for (const mode of ['subtree', 'self', 'shadow'] as const) {
			const shown = isolatedVisibilities(declaredTree, 1, mode);
			expect(shown[4], mode).toBe('visible');
			expect(shown[5], mode).toBe('visible');
		}
	});

	it('keeps an inline-visible grandchild under a child the capture hides', () => {
		expect(isolatedVisibilities(declaredTree, 1, 'self')).toEqual([
			'hidden',
			'visible',
			'hidden',
			'visible',
			'visible',
			'visible',
		]);
	});

	it('lets the capture override what the stage, target and target children declare', () => {
		const declared: IsolatedVisibilityTree = {
			parents: [-1, 0, 1],
			visibility: ['visible', 'hidden', 'visible'],
			declared: ['visible', 'hidden', 'visible'],
		};
		expect(isolatedVisibilities(declared, 1, 'self')).toEqual(['hidden', 'visible', 'hidden']);
	});
});

/** `<r ~0><a ~1>x</a>[<b ~2><c ~3/></b>]</r>`: element 2 bracketed as prunable. */
function template(overrides: Partial<IsolatedTemplate> = {}): IsolatedTemplate {
	return {
		parents: [-1, 0, 0, 2],
		visibility: ['visible', 'visible', 'visible', 'visible'],
		declared: ['', '', '', ''],
		prunable: [false, false, true, false],
		chunks: [
			'<r style="',
			'"><a style="',
			'">x</a>',
			'<b style="',
			'"><c style="',
			'"/></b>',
			'</r>',
		],
		ops: [
			{ kind: 'style', element: 0 },
			{ kind: 'style', element: 1 },
			{ kind: 'open', element: 2, end: 5 },
			{ kind: 'style', element: 2 },
			{ kind: 'style', element: 3 },
			{ kind: 'close', element: 2 },
		],
		...overrides,
	};
}

describe('isolatedMarkup', () => {
	const v = (value: string) => `visibility:${value}!important`;
	const pruned = `<r style="${v('hidden')}"><a style="${v('visible')}">x</a></r>`;

	it('writes every slot and leaves out a prunable subtree that shows nothing', () => {
		expect(isolatedMarkup(template(), 1, 'subtree')).toBe(pruned);
	});

	it('keeps the subtree when it is not pruned, or when something in it shows', () => {
		const full =
			`<r style="${v('hidden')}"><a style="${v('visible')}">x</a>` +
			`<b style="${v('hidden')}"><c style="${v('hidden')}"/></b></r>`;
		expect(isolatedMarkup(template(), 1, 'subtree', { prune: false })).toBe(full);
		expect(isolatedMarkup(template(), 3, 'subtree')).toBe(
			`<r style="${v('hidden')}"><a style="${v('hidden')}">x</a>` +
				`<b style="${v('hidden')}"><c style="${v('visible')}"/></b></r>`,
		);
	});

	it('keeps a subtree that declares itself visible, as the capture of the live stage shows it', () => {
		const shown = template({ declared: ['', '', 'visible', ''] });
		expect(isolatedMarkup(shown, 1, 'subtree')).toBe(
			`<r style="${v('hidden')}"><a style="${v('visible')}">x</a>` +
				`<b style="${v('visible')}"><c style="${v('visible')}"/></b></r>`,
		);
	});

	it('never leaves out a subtree that is not prunable', () => {
		const fixed = template({ prunable: [false, false, false, false] });
		expect(isolatedMarkup(fixed, 1, 'subtree')).toContain('<b ');
	});

	it('makes the target text, box and decorations transparent in shadow mode', () => {
		const markup = isolatedMarkup(template(), 1, 'shadow');
		const target = /<a style="([^"]*)"/u.exec(markup)?.[1] ?? '';
		expect(target.startsWith(v('visible'))).toBe(true);
		for (const declaration of [
			'color:transparent!important',
			'-webkit-text-fill-color:transparent!important',
			'background-color:transparent!important',
			'background-image:none!important',
			'border-top-color:transparent!important',
			'border-left-color:transparent!important',
			'box-shadow:none!important',
			'text-decoration-color:transparent!important',
		])
			expect(target).toContain(declaration);
		expect(/<r style="([^"]*)"/u.exec(markup)?.[1]).toBe(v('hidden'));
		expect(
			isolatedMarkup(template(), 1, 'shadow', { shadowStyle: ';color:red!important' }),
		).toContain(`<a style="${v('visible')};color:red!important"`);
	});
});
