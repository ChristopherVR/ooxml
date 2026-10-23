import { describe, expect, it } from 'vitest';

import type { LiveColorSource } from './live-color-input';
import {
	createLiveColorInput,
	createLiveColorPatchInput,
	handleLiveColorEvent,
	liveColorRevertPatch,
} from './live-color-input';

interface Shape {
	id: string;
	style: { fill: string; fillRef?: string; width?: number };
}

/**
 * A tiny editor: silent writes, committed writes that push the state from before them onto an
 * undo stack, and a recent-colours list. `lag` keeps `read()` on the value from before the
 * last silent writes, the way a binding that has not re-rendered yet still hands its controls
 * the old element.
 */
function editor(initial: Shape) {
	let shape = initial;
	let shown = initial;
	const undo: Shape[] = [];
	const recent: string[] = [];
	const silent: Partial<Shape>[] = [];
	let lag = false;
	const write = (patch: Partial<Shape>): void => {
		shape = { ...shape, ...patch };
		if (!lag) {
			shown = shape;
		}
	};
	const control = createLiveColorPatchInput<Shape, Partial<Shape>>({
		read: (): LiveColorSource<Shape> => ({ key: shown.id, color: shown.style.fill, state: shown }),
		patch: (value, hex) => ({ style: { ...value.style, fill: hex, fillRef: undefined } }),
		preview(patch) {
			silent.push(patch);
			write(patch);
		},
		commit(patch, hex) {
			undo.push(shape);
			write(patch);
			recent.unshift(hex);
		},
	});
	return {
		control,
		undo,
		recent,
		silent,
		get shape() {
			return shape;
		},
		/** An edit from elsewhere (a theme swatch, an undo): replaces the model and what is shown. */
		replace(next: Shape) {
			shape = next;
			shown = next;
		},
		/** The binding re-renders: its controls now see the live model. */
		render() {
			shown = shape;
		},
		setLag(on: boolean) {
			lag = on;
		},
	};
}

const START: Shape = { id: 'a', style: { fill: '#4472c4', fillRef: 'accent1', width: 2 } };

describe('createLiveColorPatchInput', () => {
	it('shows every shade of a drag from the start value, without an undo step or a recent colour', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.control.input('#660000');

		expect(ed.shape.style).toStrictEqual({ fill: '#660000', fillRef: undefined, width: 2 });
		expect(ed.silent).toStrictEqual([
			{ style: { fill: '#330000', fillRef: undefined, width: 2 } },
			{ style: { fill: '#660000', fillRef: undefined, width: 2 } },
		]);
		expect(ed.undo).toStrictEqual([]);
		expect(ed.recent).toStrictEqual([]);
	});

	it('commits a settled drag as one undo step back to the colour, and theme ref, from before it', () => {
		const ed = editor(START);
		for (const shade of ['#330000', '#660000', '#990000']) {
			ed.control.input(shade);
		}
		ed.control.change('#990000');

		expect(ed.undo).toStrictEqual([START]);
		expect(ed.recent).toStrictEqual(['#990000']);
		expect(ed.shape.style).toStrictEqual({ fill: '#990000', fillRef: undefined, width: 2 });
	});

	it('records nothing when the drag settles where it started, and puts the theme ref back', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.control.input('#4472C4');
		ed.control.change('#4472c4');

		expect(ed.undo).toStrictEqual([]);
		expect(ed.recent).toStrictEqual([]);
		expect(ed.shape).toStrictEqual(START);
	});

	it('commits a pick that arrives as a lone change event, and ignores one at the current colour', () => {
		const ed = editor(START);
		ed.control.change('#4472c4');
		expect(ed.undo).toStrictEqual([]);

		ed.control.change('#00ff00');
		expect(ed.undo).toStrictEqual([START]);
		expect(ed.recent).toStrictEqual(['#00ff00']);
		expect(ed.silent).toStrictEqual([]);
	});

	it('carries a drag across events while the binding still hands its controls the old value', () => {
		const ed = editor(START);
		ed.setLag(true);
		ed.control.input('#330000');
		ed.control.input('#660000');
		ed.control.change('#660000');

		expect(ed.undo).toStrictEqual([START]);
		expect(ed.shape.style.fill).toBe('#660000');
	});

	it('carries a drag across a re-render that shows an earlier shade of the same drag', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.render();
		ed.setLag(true);
		ed.control.input('#660000');
		ed.control.input('#990000');
		ed.control.change('#990000');

		expect(ed.undo).toStrictEqual([START]);
	});

	it('settles a drag on blur when no change event ended it', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.control.blur('#330000');

		expect(ed.undo).toStrictEqual([START]);
		expect(ed.recent).toStrictEqual(['#330000']);
	});

	it('takes a drag cancelled back to its start colour (Escape) out again on blur, recording nothing', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.control.input('#4472c4');
		ed.control.blur('#4472c4');

		expect(ed.undo).toStrictEqual([]);
		expect(ed.shape).toStrictEqual(START);
	});

	it('does not commit on a blur outside a drag, nor on a second blur', () => {
		const ed = editor(START);
		ed.control.blur('#123456');
		ed.control.input('#330000');
		ed.control.change('#330000');
		ed.control.blur('#330000');

		expect(ed.undo).toHaveLength(1);
		expect(ed.recent).toStrictEqual(['#330000']);
	});

	it('starts over when something else changed the colour, without writing the stale drag back', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		const themed: Shape = { id: 'a', style: { fill: '#ed7d31', fillRef: 'accent2', width: 2 } };
		ed.replace(themed);
		// A blur now is not a pick: the stale drag is dropped, the theme pick stays.
		ed.control.blur('#330000');
		expect(ed.shape).toStrictEqual(themed);

		ed.control.input('#660000');
		ed.control.change('#660000');
		expect(ed.undo).toStrictEqual([themed]);
	});

	it('starts over on another target even when it shows a colour the old drag passed', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		const other: Shape = { id: 'b', style: { fill: '#330000', width: 5 } };
		ed.replace(other);
		ed.control.input('#660000');
		ed.control.change('#4472c4');

		// The drag on `b` started at #330000: its commit records `b`, never `a`'s start state.
		expect(ed.undo).toStrictEqual([other]);
		expect(ed.shape).toStrictEqual({
			id: 'b',
			style: { fill: '#4472c4', fillRef: undefined, width: 5 },
		});
	});

	it('does nothing while there is nothing to edit', () => {
		let previews = 0;
		let commits = 0;
		const control = createLiveColorInput<null>({
			read: () => undefined,
			preview: () => {
				previews += 1;
			},
			restore: () => {
				previews += 1;
			},
			commit: () => {
				commits += 1;
			},
		});
		control.input('#ff0000');
		control.change('#ff0000');
		control.blur('#ff0000');
		expect([previews, commits]).toStrictEqual([0, 0]);
	});
});

describe('createLiveColorInput', () => {
	it('restores once from the start state before committing a settled drag', () => {
		const calls: string[] = [];
		const state = { hex: '#000000' };
		const control = createLiveColorInput<typeof state>({
			read: () => ({ key: 'k', color: state.hex, state }),
			preview: (hex, from) => calls.push(`preview ${hex} from ${from.hex}`),
			restore: (from) => calls.push(`restore ${from.hex}`),
			commit: (hex, from) => calls.push(`commit ${hex} from ${from.hex}`),
		});
		control.input('#111111');
		control.input('#222222');
		control.change('#222222');
		expect(calls).toStrictEqual([
			'preview #111111 from #000000',
			'preview #222222 from #000000',
			'restore #000000',
			'commit #222222 from #000000',
		]);
	});
});

describe('liveColorRevertPatch', () => {
	it('puts back each key as the value had it, an absent key as undefined', () => {
		expect(liveColorRevertPatch({ a: 1, b: { c: 2 } }, ['b', 'z'])).toStrictEqual({
			b: { c: 2 },
			z: undefined,
		});
		expect(liveColorRevertPatch(undefined, ['a'])).toStrictEqual({ a: undefined });
	});
});

describe('handleLiveColorEvent', () => {
	it("routes a colour input's input, change and blur events with the value it shows", () => {
		const calls: string[] = [];
		const control = {
			input: (hex: string) => calls.push(`input ${hex}`),
			change: (hex: string) => calls.push(`change ${hex}`),
			blur: (hex: string) => calls.push(`blur ${hex}`),
		};
		const fire = (type: string, value: unknown): void =>
			handleLiveColorEvent(control, { type, target: { value } } as unknown as Event);
		fire('input', ' #a1b2c3');
		fire('change', '#a1b2c3');
		fire('blur', '#a1b2c3');
		fire('focus', '#a1b2c3');
		fire('input', '');
		handleLiveColorEvent(control, { type: 'input', target: null } as unknown as Event);
		expect(calls).toStrictEqual(['input #a1b2c3', 'change #a1b2c3', 'blur #a1b2c3']);
	});
});
