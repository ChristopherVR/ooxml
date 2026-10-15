/** Generic 24px stroke glyphs for the landing's tool cards. They inherit the colour of the tile. */
const svg = (body: string): string =>
	`<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;

export const icons: Record<string, string> = {
	chat: svg('<path d="M4 5h16v11H9l-5 4z"/><path d="M8 9.5h8M8 12.5h5"/>'),
	people: svg(
		'<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/><circle cx="17" cy="9" r="2.4"/><path d="M17 14c2.4 0 4 1.8 4 4.5"/>',
	),
	video: svg('<rect x="3" y="6" width="12" height="12" rx="2"/><path d="m15 11 6-3.5v9L15 13"/>'),
	file: svg('<path d="M6 3h8l5 5v13H6z"/><path d="M14 3v5h5M9 13h7M9 17h5"/>'),
	server: svg(
		'<rect x="3" y="4" width="18" height="6" rx="1.5"/><rect x="3" y="14" width="18" height="6" rx="1.5"/><path d="M7 7h.01M7 17h.01"/>',
	),
	code: svg('<path d="m8 8-5 4 5 4M16 8l5 4-5 4M14 5l-4 14"/>'),
	arrow: svg('<path d="M7 17 17 7M9 7h8v8"/>'),
};
