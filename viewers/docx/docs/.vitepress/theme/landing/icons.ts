/**
 * Generic 24px stroke glyphs for the feature tool cards. Each value is the
 * inner markup of an <svg viewBox="0 0 24 24"> drawn with currentColor.
 */

export type IconName =
	| 'layout'
	| 'pages'
	| 'list'
	| 'comment'
	| 'shield'
	| 'file'
	| 'link'
	| 'globe';

export const ICONS: Record<IconName, string> = {
	layout: '<rect x="3.5" y="4" width="17" height="16" rx="2"/><path d="M3.5 9h17M9.5 9v11"/>',
	pages:
		'<path d="M8 3.5h8l3.5 3.5v11a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2v-12.5a2 2 0 0 1 2-2z"/><path d="M15.5 3.5V7h4"/>',
	list: '<path d="M9 6.5h11M9 12h11M9 17.5h11"/><circle cx="4.8" cy="6.5" r="1"/><circle cx="4.8" cy="12" r="1"/><circle cx="4.8" cy="17.5" r="1"/>',
	comment:
		'<path d="M4.5 5.5h15a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H12l-4.5 3.5v-3.5h-3a1 1 0 0 1-1-1v-9a1 1 0 0 1 1-1z"/>',
	shield:
		'<path d="M12 3.5l7 2.5v5.5c0 4.2-2.9 7.6-7 9-4.1-1.4-7-4.8-7-9V6z"/><path d="M9 12l2.2 2.2L15.2 10"/>',
	file: '<path d="M6 3.5h8l4 4v13H6z"/><path d="M14 3.5v4h4M9 12.5h6M9 16h6"/>',
	link: '<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>',
	globe:
		'<circle cx="12" cy="12" r="8.5"/><path d="M3.5 12h17M12 3.5c2.4 2.5 3.6 5.3 3.6 8.5s-1.2 6-3.6 8.5c-2.4-2.5-3.6-5.3-3.6-8.5S9.6 6 12 3.5z"/>',
};

export const ICON_ORDER: IconName[] = [
	'layout',
	'pages',
	'list',
	'comment',
	'shield',
	'file',
	'link',
	'globe',
];

/** The glyph markup for a feature: its own `icon`, else one picked by position. */
export function featureIcon(icon: IconName | undefined, index: number): string {
	const name = icon ?? ICON_ORDER[index % ICON_ORDER.length] ?? 'layout';
	return ICONS[name];
}
