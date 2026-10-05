// Shared foundation for the team-workspace elements. Each element is a Lit class with three
// readable parts: reactive properties, a `render()` template and a sibling `.css` file. Properties
// are declared with `declare` and set in the constructor so class fields never shadow Lit's
// accessors. Events are plain bubbling, composed CustomEvents, like every other ooxml-ui control.
import { unsafeCSS, type CSSResult } from 'lit';
import { OfficeElement, controlStyles } from '../base.js';
import resetCss from './reset.css?raw';

/** An element's style list: the shared control base, the team reset and its own .css file. */
export const withStyles = (css: string): CSSResult[] => {
	const [common, own] = controlStyles(css);
	return [common!, unsafeCSS(resetCss), own!];
};

/**
 * Base of the team-workspace elements. They keep Lit's batched (asynchronous) updates: a host
 * awaits `updateComplete`, and a chat list redraws once however many properties it sets.
 */
export class TeamsElement extends OfficeElement {
	static override syncUpdates = false;
}

/** A stable colour per id, so one person looks the same in every list. */
export function colorFor(id: string): string {
	let hue = 0;
	for (const ch of id) hue = (hue * 31 + ch.charCodeAt(0)) % 360;
	return `hsl(${hue} 45% 42%)`;
}

const DAY_MS = 86_400_000;
export function formatTime(ts: number): string {
	return new Date(ts).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}
/** "Today", "Yesterday" or a date, for the separators between days in a conversation. */
export function dayLabel(ts: number, now = Date.now()): string {
	const start = (t: number): number => new Date(t).setHours(0, 0, 0, 0);
	const diff = Math.round((start(now) - start(ts)) / DAY_MS);
	if (diff === 0) return 'Today';
	if (diff === 1) return 'Yesterday';
	return new Date(ts).toLocaleDateString([], { weekday: 'long', month: 'long', day: 'numeric' });
}
export function formatSize(bytes?: number): string {
	if (bytes === undefined) return '';
	if (bytes < 1024) return `${bytes} B`;
	if (bytes < 1_048_576) return `${Math.round(bytes / 1024)} KB`;
	return `${(bytes / 1_048_576).toFixed(1)} MB`;
}
