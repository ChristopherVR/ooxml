/** Viewports at most this tall (phone landscape) start with the ribbon collapsed to its tab row. */
export const SHORT_VIEWPORT_MEDIA = '(max-height: 500px)';

/**
 * Whether the ribbon is collapsed: an explicit choice by the user always wins, otherwise a short
 * viewport collapses it (as Office does), so the document keeps the height.
 */
export function resolveCollapsed(shortViewport: boolean, userChoice: boolean | undefined): boolean {
	return userChoice ?? shortViewport;
}
