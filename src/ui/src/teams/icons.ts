// Glyphs the team elements draw, registered in the shared ooxml-ui icon registry (20x20 stroke
// paths). `icon(name)` renders one inline so a template can place it anywhere.
import { html, svg, type TemplateResult } from 'lit';
import { getIcon, registerIcon } from '../icons';

const GLYPHS: Record<string, string> = {
	send: 'M3 10 17 3l-4 14-3-6Zm7 1 7-8',
	paperclip:
		'M15 9.5l-5.5 5.5a3.5 3.5 0 0 1-5-5l6-6a2.3 2.3 0 0 1 3.3 3.3l-6 6a1.2 1.2 0 0 1-1.7-1.7L11 6',
	smile: 'M3 10a7 7 0 1 0 14 0 7 7 0 1 0-14 0M7.5 8v.5M12.5 8v.5M7 12c1.5 2 4.5 2 6 0',
	reply: 'M8 4 3 9l5 5M3 9h9a5 5 0 0 1 5 5v2',
	hash: 'M7.5 3 6 17M14 3l-1.5 14M4 7.5h13M3 12.5h13',
	users:
		'M7 9a2.5 2.5 0 1 0 0-5 2.5 2.5 0 0 0 0 5ZM2.5 16c0-2.5 2-4 4.5-4s4.5 1.5 4.5 4M13 8.5a2 2 0 1 0 0-4M14 12c2 .3 3.5 1.6 3.5 4',
	phone:
		'M5 3h3l1.5 4-2 1.2a8 8 0 0 0 4.3 4.3L13 10.5l4 1.5v3a2 2 0 0 1-2 2A12 12 0 0 1 3 5a2 2 0 0 1 2-2Z',
	folder: 'M3 5.5h5l1.5 2H17v8H3Z',
	calendar: 'M3.5 5h13v11.5h-13ZM3.5 8.5h13M7 3v3M13 3v3',
	more: 'M5 10h.01M10 10h.01M15 10h.01',
	mic: 'M10 3a2.5 2.5 0 0 0-2.5 2.5v4a2.5 2.5 0 0 0 5 0v-4A2.5 2.5 0 0 0 10 3ZM5 9.5a5 5 0 0 0 10 0M10 14.5V17',
	micOff:
		'M10 3a2.5 2.5 0 0 0-2.5 2.5v4a2.5 2.5 0 0 0 5 0v-4A2.5 2.5 0 0 0 10 3ZM5 9.5a5 5 0 0 0 10 0M10 14.5V17M3 3l14 14',
	video: 'M3 6h9v8H3ZM12 9l5-3v8l-5-3',
	videoOff: 'M3 6h9v8H3ZM12 9l5-3v8l-5-3M3 3l14 14',
	screenShare: 'M3 4h14v9H3ZM7 17h6M10 13v4M10 10V6M8 8l2-2 2 2',
	hand: 'M7 10V4.5a1 1 0 0 1 2 0V9M9 9V3.5a1 1 0 0 1 2 0V9M11 9V4.5a1 1 0 0 1 2 0V10M13 8.5a1 1 0 0 1 2 0V12a5 5 0 0 1-5 5h-.5A4.5 4.5 0 0 1 5 13.5L4 10.5a1 1 0 0 1 2-.5l1 2',
	phoneOff: 'M3 11c4-3 10-3 14 0l-1.5 2.5-3-1V10a9 9 0 0 0-5 0v2.5l-3 1Z',
	chat: 'M3 4h14v9H8l-5 4z',
};
for (const [name, d] of Object.entries(GLYPHS)) registerIcon(name, d);

/** An inline 20x20 icon by registry name (built-in or registered above); empty when unknown. */
export function icon(name: string): TemplateResult | '' {
	const found = getIcon(name);
	if (!found) return '';
	return html`<svg class="icon" viewBox=${found.viewBox ?? '0 0 20 20'} aria-hidden="true">
		${svg`<path d=${found.d}></path>`}
	</svg>`;
}
