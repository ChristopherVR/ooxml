/**
 * Whether this viewer draws the deck's linked web content (the shared
 * `resolveExternalContentView(...).allowed`). Every surface that draws the
 * model reads it through {@link useLinkedContentView}, which withholds the
 * linked addresses until the viewer allows them; Enable Content changes this
 * value and the surfaces redraw. Nothing reloads and the model never changes.
 *
 * `PowerPointViewer` provides it for the whole viewer; the building blocks a
 * host shell renders (`SlideCanvas`, `SlidesPane`, `SlideThumbnail`) take it as
 * their `linkedContent` prop. Outside both, linked content is withheld.
 */
import { linkedContentUrl, linkedContentView } from 'ooxml-ui/pptx';
import React, { createContext, useContext } from 'react';

export const LinkedContentContext = /* @__PURE__ */ createContext(false);

/** `value` as the surfaces under this context may draw it. */
export function useLinkedContentView<T>(value: T): T {
	return linkedContentView(value, useContext(LinkedContentContext));
}

/**
 * One address a preview draws (an inspector thumbnail), withheld while it is a web address this
 * viewer has not allowed. Editing code keeps reading the model itself.
 */
export function useLinkedContentUrl(url: string | undefined): string | undefined {
	return linkedContentUrl(url, useContext(LinkedContentContext));
}

/** Provide `allowed` to `children`, or leave the surrounding context as it is when unset. */
export function LinkedContentScope({
	allowed,
	children,
}: {
	allowed: boolean | undefined;
	children: React.ReactNode;
}): React.ReactNode {
	return allowed === undefined ? (
		children
	) : (
		<LinkedContentContext.Provider value={allowed}>{children}</LinkedContentContext.Provider>
	);
}
