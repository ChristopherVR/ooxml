/**
 * Live preview for a native colour input (`<input type="color">`), decided once for every
 * binding.
 *
 * The picker fires `input` for every shade a drag passes and `change` once the pick settles
 * (and in some browsers no `change` at all when the pick ends where it started, or on Escape).
 * PowerPoint shows each shade on the slide while the user drags and makes the settled pick ONE
 * undo step and ONE Recent Colors entry. A binding that applied every shade through its normal
 * edit path recorded one undo step per shade; one that applied only on `change` showed nothing
 * until the picker closed. This controller is the rule between the two:
 *
 * - `input`: show the shade live, with no undo step, dirty flag or recent colour;
 * - `change`, or `blur` for a drag no `change` ended: put back what the previews changed, then
 *   commit the settled colour through the binding's normal edit path, so its one undo step
 *   returns to the colour from before the drag, and add it to the recent colours. A pick that
 *   settles where the drag started is no edit at all: the previews are put back (theme ref
 *   included) and no undo step is recorded. A `change` there is still the user's pick, so its
 *   colour goes into the recent colours like any other; a `blur` there is not a pick (an
 *   Escape, a drag back to the start), so it records nothing.
 *
 * The previews own only the leaves they wrote (`shapeStyle.fillColor`, its theme ref, ...),
 * each remembered with the value it had before the drag. Putting a drag back and committing
 * its pick are both rebased on the target as it is NOW: only those leaves go back, and the
 * commit is built on the result, so an edit another hand made to the same element during the
 * drag (a collaborator's line width, an assistant's text) survives, and the undo step returns
 * to a state that existed.
 *
 * A drag belongs to one target (`key`, the element). It carries over to the next event only
 * while the control edits that target and the model still shows the drag's work: the value it
 * last read (a binding that has not re-rendered since), or a colour the drag itself put there.
 * When the control has moved to another target, the drag is put back on its own target (the
 * binding's writes are addressed by key) and the event that ends it is dropped: a pick made in
 * a picker opened for one element never lands on another. When the target is the same but its
 * colour was changed by another hand (a theme swatch, an undo), the drag is dropped as it is,
 * since putting it back would overwrite that newer colour.
 *
 * @module render/live-color-input
 */

import type { LiveColorLeaf } from './live-color-leaves';
import { liveColorRevertPatch, recordLiveColorLeaves } from './live-color-leaves';

export { liveColorRevertPatch } from './live-color-leaves';
export type { LiveColorLeaf } from './live-color-leaves';

/** What a colour control edits at the moment of an event. */
export interface LiveColorSource<T> {
	/** The edited target's identity (the element id). A drag never crosses targets. */
	readonly key: string;
	/** The colour the control shows now. */
	readonly color: string;
	/** The target as the binding holds it now: the element, or the style a field reads. */
	readonly value: T;
}

/** What a binding gives {@link createLiveColorInput}; called afresh on every event. */
export interface LiveColorInputOptions<T, P extends object> {
	/** The control's target now, or `undefined` when nothing can be edited. */
	read(): LiveColorSource<T> | undefined;
	/** The patch that sets `hex` on `value`; `undefined` when `value` cannot take a colour. */
	patch(value: T, hex: string): P | undefined;
	/**
	 * Merge `patch` into the target `key` live: no undo step, dirty flag or recent colour. `key`
	 * can name a target the control has left, when the drag started there is put back.
	 */
	preview(patch: P, key: string): void;
	/**
	 * Merge the settled pick into the current target `key` through the normal edit path: ONE
	 * undo step and the dirty flag. The drag is already put back and `patch` is built on it.
	 */
	commit(patch: P, key: string, hex: string): void;
	/**
	 * Add `hex` to the recent colours: with every commit (just before it), and for a `change`
	 * that settles on the colour the drag started from (a pick, though not an edit).
	 */
	remember(hex: string): void;
}

/** The native events of one colour input, each with the input's current value. */
export interface LiveColorInput {
	/** An `input` event: one shade of a drag. */
	input(hex: string): void;
	/** A `change` event: the pick settled. */
	change(hex: string): void;
	/** A `blur` event: settles a drag no `change` ended, at the value the input shows. */
	blur(hex: string): void;
}

interface Drag<T> {
	readonly key: string;
	readonly start: string;
	/** The start colour and every shade previewed since, normalised. */
	readonly colors: Set<string>;
	readonly leaves: Map<string, LiveColorLeaf>;
	/** The target as last read: what the drag is put back on once the control has left it. */
	last: T;
}

function normalizeColor(color: string): string {
	return color.trim().toLowerCase();
}

/** Build the controller for one colour input. */
export function createLiveColorInput<T, P extends object>(
	options: LiveColorInputOptions<T, P>,
): LiveColorInput {
	let drag: Drag<T> | undefined;

	/** The drag `source` carries on, or `undefined` when it starts a new one. */
	const carriedOn = (source: LiveColorSource<T>): Drag<T> | undefined =>
		drag &&
		drag.key === source.key &&
		(source.value === drag.last || drag.colors.has(normalizeColor(source.color)))
			? drag
			: undefined;

	/**
	 * End a drag that does not carry on to `source`. One the control has left is put back on its
	 * own target and reported, so the event that found it is dropped; one whose own target was
	 * recoloured by another hand is dropped as it is.
	 */
	const leave = (source: LiveColorSource<T> | undefined): boolean => {
		const stale = drag;
		drag = undefined;
		if (!stale || (source && stale.key === source.key)) {
			return false;
		}
		if (stale.leaves.size > 0) {
			options.preview(liveColorRevertPatch<P>(stale.last, stale.leaves.values()), stale.key);
		}
		return true;
	};

	/** End the drag on `hex`; `picked` is true for a `change`, false for a `blur`. */
	const settle = (hex: string, picked: boolean): void => {
		const source = options.read();
		const current = source ? carriedOn(source) : undefined;
		// A pick for a target the control has left is dropped; a blur outside a live drag is not
		// a pick at all.
		if (!current && (leave(source) || !picked)) {
			return;
		}
		if (!source) {
			return;
		}
		drag = undefined;
		let base = source.value;
		if (current && current.leaves.size > 0) {
			const revert = liveColorRevertPatch<P>(source.value, current.leaves.values());
			options.preview(revert, current.key);
			base = { ...(source.value as object), ...revert } as unknown as T;
		}
		const start = current ? current.start : source.color;
		const moved = normalizeColor(hex) !== normalizeColor(start);
		// Recent colours first: a binding whose undo snapshot holds the list then keeps the pick
		// when the edit is undone.
		if (moved || picked) {
			options.remember(hex);
		}
		const patch = moved ? options.patch(base, hex) : undefined;
		if (patch) {
			options.commit(patch, source.key, hex);
		}
	};

	return {
		input(hex) {
			const source = options.read();
			let current = source ? carriedOn(source) : undefined;
			if (!source || !current) {
				leave(source);
			}
			if (!source) {
				return;
			}
			current ??= {
				key: source.key,
				start: source.color,
				colors: new Set([normalizeColor(source.color)]),
				leaves: new Map(),
				last: source.value,
			};
			drag = current;
			current.last = source.value;
			current.colors.add(normalizeColor(hex));
			const patch = options.patch(source.value, hex);
			if (patch) {
				recordLiveColorLeaves(current.leaves, source.value, patch);
				options.preview(patch, current.key);
			}
		},
		change: (hex) => settle(hex, true),
		blur(hex) {
			if (drag) {
				settle(hex, false);
			}
		},
	};
}

/**
 * Route a native colour input's DOM `input`, `change` or `blur` event to `control`, with the
 * value the input shows. Any other event, or an input without a value, is ignored.
 */
export function handleLiveColorEvent(control: LiveColorInput, event: Event): void {
	const target = event.target as { value?: unknown } | null;
	const hex = typeof target?.value === 'string' ? target.value.trim() : '';
	if (!hex) {
		return;
	}
	if (event.type === 'input') {
		control.input(hex);
	} else if (event.type === 'change') {
		control.change(hex);
	} else if (event.type === 'blur') {
		control.blur(hex);
	}
}
