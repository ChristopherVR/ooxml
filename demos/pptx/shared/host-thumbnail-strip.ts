/**
 * The host-owned thumbnail strip example: with `?hostStrip=1`, each custom-shell
 * demo replaces the binding's slides pane with its own strip of `SlideThumbnail`
 * building blocks, fed from the same props the pane takes. Every demo uses these
 * names, so one browser spec drives all five.
 */

/** Whether the page asked for the host-owned strip instead of the slides pane. */
export const hostThumbnailStripRequested = (): boolean =>
	new URLSearchParams(location.search).get('hostStrip') === '1';

/** Thumbnail width in CSS px; the height follows the deck's aspect ratio. */
export const HOST_STRIP_THUMBNAIL_WIDTH = 160;

/** Accessible name of the strip's navigation landmark. */
export const HOST_STRIP_LABEL = 'Host slide strip';

/** Accessible name of the button that shows slide `index` (0-based). */
export const hostStripItemLabel = (index: number): string => `Show slide ${index + 1}`;

/** The strip: a scrolling column beside the canvas. */
export const HOST_STRIP_STYLE: Readonly<Record<string, string>> = {
	flex: 'none',
	display: 'flex',
	flexDirection: 'column',
	gap: '8px',
	padding: '8px',
	overflowY: 'auto',
	width: `${HOST_STRIP_THUMBNAIL_WIDTH + 16}px`,
	boxSizing: 'border-box',
	background: '#cbd5e1',
};

/** One strip button; the current slide gets a thicker outline. */
export function hostStripItemStyle(active: boolean): Readonly<Record<string, string>> {
	return {
		display: 'block',
		flex: 'none',
		width: `${HOST_STRIP_THUMBNAIL_WIDTH}px`,
		padding: '0',
		border: '0',
		background: '#fff',
		cursor: 'pointer',
		outline: active ? '3px solid #2563eb' : '1px solid #94a3b8',
	};
}

/** A style map as a `style` attribute value, for templates that bind strings. */
export function cssText(style: Readonly<Record<string, string>>): string {
	return Object.entries(style)
		.map(([key, value]) => `${key.replace(/[A-Z]/gu, (c) => `-${c.toLowerCase()}`)}:${value}`)
		.join(';');
}
