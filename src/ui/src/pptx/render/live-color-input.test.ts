import { describe, expect, it } from 'vitest';

import type { LiveColorSource } from './live-color-input';
import { createLiveColorInput, handleLiveColorEvent } from './live-color-input';

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
	const control = createLiveColorInput<Shape, Partial<Shape>>({
		read: (): LiveColorSource<Shape> => ({ key: shown.id, color: shown.style.fill, value: shown }),
		patch: (value, hex) => ({ style: { ...value.style, fill: hex, fillRef: undefined } }),
		lookup: (key) => (key === shape.id ? shape : undefined),
		preview(patch) {
			silent.push(patch);
			write(patch);
		},
		commit(patch) {
			undo.push(shape);
			write(patch);
		},
		remember(hex) {
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

/** Two shapes and a selection, for drags the selection leaves: writes go to the shape named. */
function deck() {
	const shapes = new Map<string, Shape>([
		['a', START],
		['b', { id: 'b', style: { fill: '#00aa55', fillRef: 'accent6', width: 7 } }],
	]);
	let selected = 'a';
	const undo: Shape[] = [];
	const recent: string[] = [];
	// Like a binding's write by id: nothing happens for an element that is gone.
	const write = (patch: Partial<Shape>, key: string): void => {
		const shape = shapes.get(key);
		if (shape) {
			shapes.set(key, { ...shape, ...patch });
		}
	};
	const control = createLiveColorInput<Shape, Partial<Shape>>({
		read: () => {
			const shape = shapes.get(selected)!;
			return { key: shape.id, color: shape.style.fill, value: shape };
		},
		patch: (value, hex) => ({ style: { ...value.style, fill: hex, fillRef: undefined } }),
		lookup: (key) => shapes.get(key),
		preview: (patch, key) => write(patch, key),
		commit(patch, key) {
			undo.push(shapes.get(key)!);
			write(patch, key);
		},
		remember: (hex) => recent.unshift(hex),
	});
	return {
		control,
		undo,
		recent,
		shape: (id: string) => shapes.get(id),
		select(id: string) {
			selected = id;
		},
		/** Another hand edits a shape's style (a collaborator, the assistant). */
		edit(id: string, style: Partial<Shape['style']>) {
			const shape = shapes.get(id)!;
			shapes.set(id, { ...shape, style: { ...shape.style, ...style } });
		},
		remove(id: string) {
			shapes.delete(id);
		},
	};
}

describe('createLiveColorInput', () => {
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

	it('makes no undo step when the drag settles where it started, and puts the theme ref back', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.control.input('#4472C4');
		ed.control.change('#4472c4');

		expect(ed.undo).toStrictEqual([]);
		expect(ed.shape).toStrictEqual(START);
		// The `change` is still the user's pick, so it counts as a recent colour.
		expect(ed.recent).toStrictEqual(['#4472c4']);
	});

	it('commits a pick that arrives as a lone change event; one at the current colour is no edit', () => {
		const ed = editor(START);
		ed.control.change('#4472c4');
		expect(ed.undo).toStrictEqual([]);
		expect(ed.recent).toStrictEqual(['#4472c4']);

		ed.control.change('#00ff00');
		expect(ed.undo).toStrictEqual([START]);
		expect(ed.recent).toStrictEqual(['#00ff00', '#4472c4']);
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
		expect(ed.recent).toStrictEqual([]);
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

	it('does nothing while there is nothing to edit', () => {
		let previews = 0;
		let commits = 0;
		const control = createLiveColorInput<null, object>({
			read: () => undefined,
			patch: () => ({}),
			lookup: () => undefined,
			preview: () => {
				previews += 1;
			},
			commit: () => {
				commits += 1;
			},
			remember: () => {
				commits += 1;
			},
		});
		control.input('#ff0000');
		control.change('#ff0000');
		control.blur('#ff0000');
		expect([previews, commits]).toStrictEqual([0, 0]);
	});
});

describe('a drag the selection leaves', () => {
	it('is put back on its own target, and the pick that ends it touches nothing else', () => {
		const d = deck();
		const b = d.shape('b');
		d.control.input('#330000');
		d.select('b');
		d.control.change('#330000');

		expect(d.shape('a')).toStrictEqual(START);
		expect(d.shape('b')).toStrictEqual(b);
		expect(d.undo).toStrictEqual([]);
		expect(d.recent).toStrictEqual([]);
	});

	it('is put back when a drag starts on the new target, which previews on its own', () => {
		const d = deck();
		d.control.input('#330000');
		d.select('b');
		d.control.input('#660000');

		expect(d.shape('a')).toStrictEqual(START);
		expect(d.shape('b')?.style.fill).toBe('#660000');
	});

	it('starts over on the new target even when it shows a colour the old drag passed', () => {
		const d = deck();
		d.control.input('#330000');
		d.control.input('#00aa55');
		d.select('b');
		d.control.input('#660000');
		d.control.change('#123456');

		// The drag on `b` started at its own colour: its one undo step records `b`, never `a`.
		expect(d.shape('a')).toStrictEqual(START);
		expect(d.undo).toStrictEqual([
			{ id: 'b', style: { fill: '#00aa55', fillRef: 'accent6', width: 7 } },
		]);
		expect(d.shape('b')?.style).toStrictEqual({ fill: '#123456', fillRef: undefined, width: 7 });
	});

	it('is put back on the first target as it is now, keeping an edit made there after its last shade', () => {
		const d = deck();
		d.control.input('#330000');
		d.edit('a', { width: 5 });
		d.select('b');
		d.control.change('#330000');

		expect(d.shape('a')).toStrictEqual({ id: 'a', style: { ...START.style, width: 5 } });
	});

	it('writes nothing when the first target is gone', () => {
		const d = deck();
		const b = d.shape('b');
		d.control.input('#330000');
		d.remove('a');
		d.select('b');
		d.control.change('#330000');

		expect(d.shape('a')).toBeUndefined();
		expect(d.shape('b')).toStrictEqual(b);
	});

	it('is put back on blur too, without committing to the new target', () => {
		const d = deck();
		const b = d.shape('b');
		d.control.input('#330000');
		d.select('b');
		d.control.blur('#330000');

		expect(d.shape('a')).toStrictEqual(START);
		expect(d.shape('b')).toStrictEqual(b);
		expect(d.undo).toStrictEqual([]);
	});
});

describe('another edit to the same element while a drag runs', () => {
	// A collaborator, an undo or the assistant changes a field next to the colour mid-drag.
	it('survives the commit, and the undo step returns to the colour with that edit kept', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.replace({ ...ed.shape, style: { ...ed.shape.style, width: 5 } });
		ed.control.change('#330000');

		expect(ed.shape.style).toStrictEqual({ fill: '#330000', fillRef: undefined, width: 5 });
		expect(ed.undo).toStrictEqual([{ id: 'a', style: { ...START.style, width: 5 } }]);
	});

	it('survives a drag that settles where it started', () => {
		const ed = editor(START);
		ed.control.input('#330000');
		ed.replace({ ...ed.shape, style: { ...ed.shape.style, width: 5 } });
		ed.control.change('#4472c4');

		expect(ed.shape).toStrictEqual({ id: 'a', style: { ...START.style, width: 5 } });
		expect(ed.undo).toStrictEqual([]);
	});
});

describe('a preview that creates an entry the target did not have', () => {
	// A text colour on an element with no `textStyle`: the preview creates the whole object.
	function bare() {
		let value: { id: string; style?: Record<string, unknown> } = { id: 'a' };
		const control = createLiveColorInput<typeof value, Partial<typeof value>>({
			read: () => ({ key: 'a', color: String(value.style?.['fill'] ?? '#000000'), value }),
			patch: (current, hex) => ({ style: { ...current.style, fill: hex } }),
			lookup: () => value,
			preview(patch) {
				value = { ...value, ...patch };
			},
			commit: () => {},
			remember: () => {},
		});
		return {
			control,
			get value() {
				return value;
			},
			edit(style: Record<string, unknown>) {
				value = { ...value, style: { ...value.style, ...style } };
			},
		};
	}

	it('puts back only what it added, keeping what another edit put inside since', () => {
		const t = bare();
		t.control.input('#330000');
		t.edit({ bold: true });
		t.control.input('#000000');
		t.control.blur('#000000');

		expect(t.value).toStrictEqual({ id: 'a', style: { bold: true } });
	});

	it('takes the entry away again when nothing else went into it', () => {
		const t = bare();
		t.control.input('#330000');
		t.control.input('#000000');
		t.control.blur('#000000');

		expect(t.value.style).toBeUndefined();
	});
});

describe('putting a drag back', () => {
	it('removes the entries its previews added, an emptied theme ref included', () => {
		let value: { style: Record<string, unknown> } = { style: { fill: '#000000' } };
		const control = createLiveColorInput<typeof value, Partial<typeof value>>({
			read: () => ({ key: 'k', color: String(value.style['fill']), value }),
			patch: (current, hex) => ({
				style: { ...current.style, fill: hex, ref: undefined, width: 1 },
			}),
			lookup: () => value,
			preview(patch) {
				value = { ...value, ...patch };
			},
			commit: () => {},
			remember: () => {},
		});
		control.input('#330000');
		control.input('#000000');
		control.blur('#000000');

		expect(value).toStrictEqual({ style: { fill: '#000000' } });
	});
});

describe('a pick whose input and change arrive in the same turn', () => {
	// What a picker that commits at once does, and what Playwright's `fill()` does: no render
	// runs between the two events, so a lagging binding still hands over the old value.
	it.each([false, true])('commits a new colour once (binding lagging: %s)', (lag) => {
		const ed = editor(START);
		ed.setLag(lag);
		ed.control.input('#00ff00');
		ed.control.change('#00ff00');

		expect(ed.undo).toStrictEqual([START]);
		expect(ed.recent).toStrictEqual(['#00ff00']);
		expect(ed.shape.style).toStrictEqual({ fill: '#00ff00', fillRef: undefined, width: 2 });
	});

	it.each([false, true])(
		'counts a pick of the colour already there as a recent colour, with no undo step (binding lagging: %s)',
		(lag) => {
			const ed = editor(START);
			ed.setLag(lag);
			ed.control.input('#4472c4');
			ed.control.change('#4472c4');

			expect(ed.recent).toStrictEqual(['#4472c4']);
			expect(ed.undo).toStrictEqual([]);
			expect(ed.shape).toStrictEqual(START);
		},
	);
});

describe('the writes a drag makes', () => {
	it('previews, puts back, remembers, then commits, each addressed to the target', () => {
		const calls: string[] = [];
		let value = { key: 'k', fill: '#000000', width: 1 };
		const control = createLiveColorInput<typeof value, Partial<typeof value>>({
			read: () => ({ key: value.key, color: value.fill, value }),
			patch: (_value, hex) => ({ fill: hex }),
			lookup: () => value,
			preview(patch, key) {
				calls.push(`preview ${JSON.stringify(patch)} to ${key}`);
				value = { ...value, ...patch };
			},
			commit: (patch, key, hex) => calls.push(`commit ${JSON.stringify(patch)} to ${key} (${hex})`),
			remember: (hex) => calls.push(`remember ${hex}`),
		});
		control.input('#111111');
		control.input('#222222');
		control.change('#222222');
		expect(calls).toStrictEqual([
			'preview {"fill":"#111111"} to k',
			'preview {"fill":"#222222"} to k',
			'preview {"fill":"#000000"} to k',
			'remember #222222',
			'commit {"fill":"#222222"} to k (#222222)',
		]);
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
