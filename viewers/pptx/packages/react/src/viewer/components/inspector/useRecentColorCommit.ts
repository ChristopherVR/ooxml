import { useCallback, useEffect, useRef } from 'react';

import { useRecentColors } from './RecentColorsContext';

/** The latest `pushColor` of the hook that last attached each input. */
const pushers = new WeakMap<HTMLInputElement, { current: (hex: string) => void }>();

/** One listener for every input, so attaching an input again adds nothing. */
function pushCommittedColor(event: Event): void {
	const input = event.currentTarget as HTMLInputElement;
	pushers.get(input)?.current(input.value);
}

/**
 * A ref callback for a native `<input type="color">` that adds the picked
 * colour to the deck's "Recent Colors" list (`p:clrMru`) once the pick is
 * COMMITTED.
 *
 * React's `onChange` on an input runs for every native `input` event, which a
 * colour input fires continuously while the user drags through the OS picker,
 * so pushing from `onChange` filled the capped list with every shade passed on
 * the way. The native `change` event fires once, when the picker settles on a
 * value, and React does not expose it separately, so this listens for it
 * directly. Keep the live apply in `onChange` and pass this as the input's
 * `ref`:
 *
 *     const commitRecentColor = useRecentColorCommit();
 *     <input type='color' ref={commitRecentColor} onChange={(e) => apply(e.target.value)} />
 *
 * One callback can serve several inputs (a list of gradient stops), and it
 * always pushes through the latest `pushColor`. It returns no cleanup, so it
 * behaves the same on React 18 (whose ref callbacks cannot return one) and 19:
 * the listener goes away with its element, and since every input shares one
 * listener function, `addEventListener` ignores a repeated attach (StrictMode,
 * a new callback), so a pick is pushed once.
 */
export function useRecentColorCommit(): (input: HTMLInputElement | null) => void {
	const { pushColor } = useRecentColors();
	const pushRef = useRef(pushColor);
	useEffect(() => {
		pushRef.current = pushColor;
	}, [pushColor]);
	return useCallback((input: HTMLInputElement | null) => {
		if (input) {
			pushers.set(input, pushRef);
			input.addEventListener('change', pushCommittedColor);
		}
	}, []);
}
