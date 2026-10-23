import { createLiveColorInput } from 'ooxml-ui/pptx';
import type { LiveColorInput as LiveColorController } from 'ooxml-ui/pptx';
import type { PptxThemeColorRef } from 'pptx-viewer-core';
import React, { useCallback, useRef } from 'react';

import { useLiveEdit } from './LiveEditContext';
import { useRecentColors } from './RecentColorsContext';

/** The controller of the component that last attached each input. */
const controllers = new WeakMap<HTMLInputElement, { current: LiveColorController | null }>();

/** One native `change` listener for every input, so attaching an input again adds nothing. */
function settleOnChange(event: Event): void {
	const input = event.currentTarget as HTMLInputElement;
	controllers.get(input)?.current?.change(input.value);
}

interface Latest {
	value: string;
	selectedRef: PptxThemeColorRef | undefined;
	onApply: (hex: string, ref?: PptxThemeColorRef) => void;
	withoutUndoStep: (apply: () => void) => void;
	pushColor: (hex: string) => void;
}

/** The input's controller, reading the latest props through `latest` on every event. */
function liveController(latest: { current: Latest }): LiveColorController {
	return createLiveColorInput<{ hex: string; ref: PptxThemeColorRef | undefined }>({
		read: () => ({
			key: 'color',
			color: latest.current.value,
			state: { hex: latest.current.value, ref: latest.current.selectedRef },
		}),
		preview: (hex) => latest.current.withoutUndoStep(() => latest.current.onApply(hex)),
		restore: (from) =>
			latest.current.withoutUndoStep(() => latest.current.onApply(from.hex, from.ref)),
		commit: (hex) => latest.current.onApply(hex),
		remember: (hex) => latest.current.pushColor(hex),
	});
}

export interface LiveColorInputProps extends Omit<
	React.InputHTMLAttributes<HTMLInputElement>,
	'type' | 'value' | 'onChange' | 'onBlur'
> {
	/** The colour the model holds now. */
	value: string;
	/** The theme ref `value` came from, put back when a drag settles where it started. */
	selectedRef?: PptxThemeColorRef;
	/**
	 * Apply a colour through the control's normal edit path: each shade of a
	 * drag (inside the viewer's `withoutUndoStep`, so it is no undo step), the
	 * start colour and ref when a drag is put back, and the settled pick.
	 */
	onApply: (hex: string, ref?: PptxThemeColorRef) => void;
}

/**
 * A native `<input type="color">` that previews while the user drags through
 * the picker and commits once when the pick settles (shared
 * `createLiveColorInput`): every shade shows on the slide with no undo step or
 * recent colour, the settled pick (native `change`, or `blur` for a drag no
 * change ended) is ONE undo step and one Recent Colors entry, and a drag that
 * settles where it started is no undo step (its `change` still counts as a
 * recent colour).
 *
 * React's `onChange` runs for every native `input`, and React does not expose
 * the native `change` on its own, so that one is listened for directly, the
 * way `useRecentColorCommit` does.
 */
export function LiveColorInput({
	value,
	selectedRef,
	onApply,
	...rest
}: LiveColorInputProps): React.ReactElement {
	const { withoutUndoStep } = useLiveEdit();
	const { pushColor } = useRecentColors();
	const latest = useRef<Latest>({ value, selectedRef, onApply, withoutUndoStep, pushColor });
	latest.current = { value, selectedRef, onApply, withoutUndoStep, pushColor };
	// Created once per component: a drag's state must survive the re-render each shade causes.
	const controlRef = useRef<LiveColorController | null>(null);
	controlRef.current ??= liveController(latest);
	const control = controlRef.current;
	const attach = useCallback((input: HTMLInputElement | null) => {
		if (input) {
			controllers.set(input, controlRef);
			input.addEventListener('change', settleOnChange);
		}
	}, []);
	return (
		<input
			{...rest}
			type='color'
			value={value}
			ref={attach}
			onChange={(event) => control.input(event.target.value)}
			onBlur={(event) => control.blur(event.target.value)}
		/>
	);
}
