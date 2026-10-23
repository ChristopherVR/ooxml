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
 *   returns to the colour from before the drag. A pick that settles where the drag started is
 *   no edit at all: the previews are put back and nothing is recorded.
 *
 * A drag belongs to one target. It carries over to the next event only while the control edits
 * the same `key` and the model still shows the drag's work: the untouched start value (a
 * binding that has not re-rendered since), or a colour the drag itself put there. Anything
 * else (another selection, an undo, a theme swatch pick) starts a new drag from what the model
 * holds now, so a stale drag can never write its start state over a newer edit.
 *
 * @module render/live-color-input
 */

/** What a colour control edits at the moment of an event. */
export interface LiveColorSource<S> {
	/** The edited object's identity (the element id, or the field id for a style field). */
	readonly key: string;
	/** The colour the control shows now. */
	readonly color: string;
	/** What a drag needs to preview from and to put back: the element, or the style it reads. */
	readonly state: S;
}

/** What a binding gives {@link createLiveColorInput}; called afresh on every event. */
export interface LiveColorInputOptions<S> {
	/** The control's target now, or `undefined` when nothing can be edited. */
	read(): LiveColorSource<S> | undefined;
	/** Show `hex` over `from` (the drag's start) without an undo step, dirty flag or recent colour. */
	preview(hex: string, from: S): void;
	/** Put back, again silently, everything the drag's previews changed on `from`. */
	restore(from: S): void;
	/** Commit `hex` over `from` as ONE undo step, mark the deck dirty and record the recent colour. */
	commit(hex: string, from: S): void;
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

interface Drag<S> {
	readonly key: string;
	readonly start: string;
	readonly state: S;
	/** The start colour and every shade previewed since, normalised. */
	readonly colors: Set<string>;
}

function normalizeColor(color: string): string {
	return color.trim().toLowerCase();
}

/** Build the controller for one colour input. */
export function createLiveColorInput<S>(options: LiveColorInputOptions<S>): LiveColorInput {
	let drag: Drag<S> | undefined;

	/** The drag `source` carries on, or `undefined` when it starts a new one. */
	const carriedOn = (source: LiveColorSource<S>): Drag<S> | undefined =>
		drag &&
		drag.key === source.key &&
		(source.state === drag.state || drag.colors.has(normalizeColor(source.color)))
			? drag
			: undefined;

	/**
	 * End the drag on `hex`. A drag that is still carried on is put back first, so the commit
	 * records the state from before it; a stale one is dropped as it is, since putting it back
	 * would overwrite whatever changed the model since.
	 */
	const settle = (source: LiveColorSource<S>, hex: string): void => {
		const current = carriedOn(source);
		drag = undefined;
		if (current) {
			options.restore(current.state);
		}
		const from = current ? current.state : source.state;
		const start = current ? current.start : source.color;
		if (normalizeColor(hex) !== normalizeColor(start)) {
			options.commit(hex, from);
		}
	};

	return {
		input(hex) {
			const source = options.read();
			if (!source) {
				return;
			}
			const current = carriedOn(source) ?? {
				key: source.key,
				start: source.color,
				state: source.state,
				colors: new Set([normalizeColor(source.color)]),
			};
			drag = current;
			current.colors.add(normalizeColor(hex));
			options.preview(hex, current.state);
		},
		change(hex) {
			const source = options.read();
			if (source) {
				settle(source, hex);
			} else {
				drag = undefined;
			}
		},
		blur(hex) {
			// Only a drag still in progress settles here: a blur is not a pick, so with no drag (or
			// a stale one) it must not commit whatever the input happens to show.
			const source = drag ? options.read() : undefined;
			if (source && carriedOn(source)) {
				settle(source, hex);
			} else {
				drag = undefined;
			}
		},
	};
}

/** What a binding gives {@link createLiveColorPatchInput}: the colour as a patch on a value. */
export interface LiveColorPatchOptions<T, P extends object> {
	/** The edited value now (the element, or the style a field reads). */
	read(): LiveColorSource<T> | undefined;
	/** The patch that sets `hex` on `value`; `undefined` when `value` cannot take a colour. */
	patch(value: T, hex: string): P | undefined;
	/** Write `patch` live: no undo step, dirty flag or recent colour. */
	preview(patch: P): void;
	/** Write `patch` through the normal edit path (one undo step) and record `hex` as recent. */
	commit(patch: P, hex: string): void;
}

/**
 * The patch that puts back every key in `keys` as `value` had it, the way a drag's previews
 * are undone before its commit. A key `value` did not have comes back as `undefined`.
 */
export function liveColorRevertPatch<P extends object>(value: unknown, keys: Iterable<string>): P {
	const source =
		typeof value === 'object' && value !== null ? (value as Record<string, unknown>) : {};
	const patch: Record<string, unknown> = {};
	for (const key of keys) {
		patch[key] = source[key];
	}
	return patch as P;
}

/**
 * {@link createLiveColorInput} for a control whose colour is a patch on a value (an element
 * or a style object): each shade is the patch from the drag's start value, and the previews
 * are put back by writing the start value's own entries for every key they touched.
 */
export function createLiveColorPatchInput<T, P extends object>(
	options: LiveColorPatchOptions<T, P>,
): LiveColorInput {
	const touched = new Set<string>();
	return createLiveColorInput<T>({
		read: () => options.read(),
		preview(hex, from) {
			const patch = options.patch(from, hex);
			if (!patch) {
				return;
			}
			for (const key of Object.keys(patch)) {
				touched.add(key);
			}
			options.preview(patch);
		},
		restore(from) {
			if (touched.size > 0) {
				options.preview(liveColorRevertPatch<P>(from, touched));
				touched.clear();
			}
		},
		commit(hex, from) {
			touched.clear();
			const patch = options.patch(from, hex);
			if (patch) {
				options.commit(patch, hex);
			}
		},
	});
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
