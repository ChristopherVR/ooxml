/**
 * The in-page fallback for copied slide packages.
 *
 * Safari and iPadOS do not accept custom `web ` clipboard formats, so a copied slide package cannot
 * travel through the system clipboard there. The package is kept here instead, at module level, so
 * every viewer in the same document (one deck to another) can still paste it. Only the latest copy
 * is kept, and nothing leaves the page.
 */

/** A slide package held in the page, with the text preview written beside it. */
export interface InPageSlidePackage {
	/** The serialized package (`serializeSlidePackageClipboard` output). */
	readonly text: string;
	/** The plain-text preview offered to other applications. */
	readonly preview: string;
}

let stored: InPageSlidePackage | null = null;

/** Keep `value` as the page's copied slide package, replacing the previous one. */
export function storeInPageSlidePackage(value: InPageSlidePackage): void {
	stored = value;
}

/** The slide package last copied in this page, or `null`. */
export function readInPageSlidePackage(): InPageSlidePackage | null {
	return stored;
}

/** Forget the page's copied slide package (a failed copy, or tests). */
export function clearInPageSlidePackage(): void {
	stored = null;
}
