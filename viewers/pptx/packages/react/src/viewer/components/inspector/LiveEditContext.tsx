/**
 * LiveEditContext: how an inspector control applies a live frame without an
 * undo step (a shade of a colour picker drag). `ViewerInspector` provides the
 * viewer history's `withoutUndoStep`, so the shades a drag passes run through
 * the control's normal edit callback but do not arm the undo tracker, and the
 * settled pick is the one undo step.
 *
 * A context rather than a prop, for the same reason as `RecentColorsContext`:
 * the colour inputs sit several components below the pane, behind callbacks
 * that only know their own field. Without a provider (an isolated component)
 * every shade is an ordinary edit, as before.
 */
import { createContext, useContext } from 'react';

export interface LiveEditContextValue {
	/** Run `apply` so that its edits show at once but are no undo step. */
	withoutUndoStep: (apply: () => void) => void;
}

const DEFAULT_VALUE: LiveEditContextValue = {
	withoutUndoStep: (apply) => apply(),
};

const LiveEditContext = /* @__PURE__ */ createContext<LiveEditContextValue>(DEFAULT_VALUE);

export const LiveEditProvider = LiveEditContext.Provider;

export function useLiveEdit(): LiveEditContextValue {
	return useContext(LiveEditContext);
}
