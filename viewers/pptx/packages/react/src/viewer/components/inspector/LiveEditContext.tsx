/**
 * LiveEditContext: how an inspector colour control shows a live frame (a
 * shade of a colour picker drag) without an undo step. `ViewerInspector`
 * provides the element the inspector edits, a silent element write (the
 * viewer's element update inside the history's `withoutUndoStep`) and a read of
 * any element as the deck holds it, so a drag previews and puts back exactly the
 * element leaves it changes (on the element as it is then, also once the
 * selection has left it), and the settled pick goes through the control's
 * normal edit callback as the one undo step.
 *
 * A context rather than a prop, for the same reason as `RecentColorsContext`:
 * the colour inputs sit several components below the pane, behind callbacks
 * that only know their own field. Without a provider (an isolated component)
 * a colour input is a plain one: every shade is an ordinary edit, as before.
 */
import type { PptxElement } from 'pptx-viewer-core';
import { createContext, useContext } from 'react';

export interface LiveEditContextValue {
	/** The element the inspector edits, as last rendered; `null` with nothing selected. */
	element: PptxElement | null;
	/**
	 * Merge `patch` into element `id` (the inspector's element, or one a drag has
	 * left) at once, with no undo step and no dirty flag.
	 */
	previewElement: (id: string, patch: Partial<PptxElement>) => void;
	/**
	 * Element `id` as the deck holds it now, or `undefined` when it is gone: what a
	 * drag the selection has left is put back on (nothing is, when it is gone).
	 */
	elementById: (id: string) => PptxElement | undefined;
}

const LiveEditContext = /* @__PURE__ */ createContext<LiveEditContextValue | null>(null);

export const LiveEditProvider = LiveEditContext.Provider;

export function useLiveEdit(): LiveEditContextValue | null {
	return useContext(LiveEditContext);
}
