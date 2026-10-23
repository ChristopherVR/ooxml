import { createLiveColorInput } from 'ooxml-ui/pptx';
import type { LiveColorInput as LiveColorController } from 'ooxml-ui/pptx';
import type { PptxElement, PptxThemeColorRef } from 'pptx-viewer-core';
import React, { useCallback, useRef } from 'react';

import type { LiveEditContextValue } from './LiveEditContext';
import { useLiveEdit } from './LiveEditContext';
import { useRecentColors } from './RecentColorsContext';

/** How the component that last attached each input handles its native `change`. */
const changeHandlers = new WeakMap<HTMLInputElement, { current: (hex: string) => void }>();

/** One native `change` listener for every input, so attaching an input again adds nothing. */
function settleOnChange(event: Event): void {
	const input = event.currentTarget as HTMLInputElement;
	changeHandlers.get(input)?.current(input.value);
}

/** The patch a colour pick makes on an element. */
export type LiveColorPatch = (el: PptxElement, hex: string) => Partial<PptxElement> | undefined;

export interface LiveColorInputProps extends Omit<
	React.InputHTMLAttributes<HTMLInputElement>,
	'type' | 'value' | 'onChange' | 'onBlur'
> {
	/** The colour the control shows now. */
	value: string;
	/** Commit a pick through the control's normal edit path (one undo step). */
	onApply: (hex: string, ref?: PptxThemeColorRef) => void;
	/**
	 * The patch the same pick makes on the inspector's element, which the drag's
	 * shades are written with (silently) and put back from, leaf by leaf. Without
	 * it, or outside the inspector, the input is a plain one: every shade goes
	 * through `onApply`.
	 */
	patch?: LiveColorPatch;
}

interface Latest {
	value: string;
	onApply: (hex: string) => void;
	patch: LiveColorPatch | undefined;
	live: LiveEditContextValue | null;
	pushColor: (hex: string) => void;
}

/** The input's controller, reading the latest props through `latest` on every event. */
function liveController(latest: { current: Latest }): LiveColorController {
	return createLiveColorInput<PptxElement, Partial<PptxElement>>({
		read: () => {
			const { live, value } = latest.current;
			return live?.element
				? { key: live.element.id, color: value, value: live.element }
				: undefined;
		},
		patch: (el, hex) => latest.current.patch?.(el, hex),
		preview: (patch, key) => latest.current.live?.previewElement(key, patch),
		commit: (_patch, _key, hex) => latest.current.onApply(hex),
		remember: (hex) => latest.current.pushColor(hex),
	});
}

/**
 * A native `<input type="color">` that previews while the user drags through
 * the picker and commits once when the pick settles (shared
 * `createLiveColorInput`): every shade shows on the slide with no undo step or
 * recent colour, the settled pick (native `change`, or `blur` for a drag no
 * change ended) is ONE undo step and one Recent Colors entry, and a drag that
 * settles where it started is no undo step (its `change` still counts as a
 * recent colour). Putting a drag back writes the element's own leaves, so an
 * element with no colour of its own keeps none. The drag belongs to the
 * element it started on: a new selection never receives its pick.
 *
 * React's `onChange` runs for every native `input`, and React does not expose
 * the native `change` on its own, so that one is listened for directly, the
 * way `useRecentColorCommit` does.
 */
export function LiveColorInput({
	value,
	onApply,
	patch,
	...rest
}: LiveColorInputProps): React.ReactElement {
	const live = useLiveEdit();
	const { pushColor } = useRecentColors();
	const latest = useRef<Latest>({ value, onApply, patch, live, pushColor });
	latest.current = { value, onApply, patch, live, pushColor };
	const isLive = Boolean(live && patch);
	// Created once per component: a drag's state must survive the re-render each shade causes.
	const controlRef = useRef<LiveColorController | null>(null);
	controlRef.current ??= liveController(latest);
	const control = controlRef.current;
	const onChangeRef = useRef((hex: string) => {
		if (latest.current.live && latest.current.patch) {
			control.change(hex);
		} else {
			latest.current.pushColor(hex);
		}
	});
	const attach = useCallback((input: HTMLInputElement | null) => {
		if (input) {
			changeHandlers.set(input, onChangeRef);
			input.addEventListener('change', settleOnChange);
		}
	}, []);
	return (
		<input
			{...rest}
			type='color'
			value={value}
			ref={attach}
			onChange={(event) => {
				if (!isLive) {
					onApply(event.target.value);
				} else if (event.nativeEvent.type === 'input') {
					// React also reports a native `change` here when the value moved since its last
					// render; that one is the settled pick, which the listener above already handles.
					control.input(event.target.value);
				}
			}}
			onBlur={(event) => {
				if (isLive) {
					control.blur(event.target.value);
				}
			}}
		/>
	);
}
