import type { RibbonIconArtwork } from './types';

/** A bold "S" (x 6.7-14.4, baseline 17), outlined for Text Shadow. */
const SHADOW_S =
	'M10.52 5.63L11.53 5.61L12.4 5.67L13.25 5.82L13.9 6.04L13.9 8.34L13.28 8.01L12.58 7.77L11.58 7.59L10.88 7.6L10.13 7.76L9.7 8.02L9.44 8.35L9.36 8.78L9.4 8.98L9.51 9.23L9.93 9.63L10.6 10.03L12.08 10.69L12.68 11.02L13.23 11.39L13.7 11.85L14.08 12.4L14.33 13.15L14.39 14.08L14.31 14.7L14.03 15.43L13.63 15.98L13.05 16.46L12.28 16.83L11.28 17.08L10.3 17.18L8.95 17.13L7.78 16.93L7.15 16.73L6.76 16.55L6.77 14.08L7.43 14.53L8.03 14.83L8.7 15.05L9.5 15.19L10.45 15.16L11 15.01L11.48 14.71L11.71 14.3L11.73 13.95L11.66 13.68L11.51 13.43L11.23 13.14L10.5 12.69L9 12.03L8.48 11.76L7.95 11.41L7.53 11.03L7.05 10.38L6.87 9.98L6.74 9.45L6.72 8.63L6.84 7.93L7.07 7.4L7.5 6.83L8 6.4L8.7 6.02L9.5 5.77Z';
/** A bold "AV" (x 4-18.4, baseline 16), outlined for Character Spacing. */
const SPACING_AV =
	'M6.84 8.3L8.91 8.3L11.63 15.99L9.75 15.98L9.18 14.28L6.48 14.28L5.92 15.98L4.04 15.98ZM11.17 8.3L13.04 8.3L14.63 13.6L14.83 14.41L14.97 13.73L16.59 8.3L18.41 8.3L15.75 15.99L13.79 15.98ZM7.83 9.66L7.68 10.3L6.85 12.94L8.8 12.93L7.97 10.33Z';

export const FONT_ICONS: Record<string, RibbonIconArtwork> = {
	'home.font.bold': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{
				tag: 'path',
				attrs: { d: 'M6 12h9a4 4 0 0 1 0 8H7a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h7a4 4 0 0 1 0 8' },
			},
		],
	},
	'home.font.italic': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{ tag: 'line', attrs: { x1: '19', x2: '10', y1: '4', y2: '4' } },
			{ tag: 'line', attrs: { x1: '14', x2: '5', y1: '20', y2: '20' } },
			{ tag: 'line', attrs: { x1: '15', x2: '9', y1: '4', y2: '20' } },
		],
	},
	'home.font.underline': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{ tag: 'path', attrs: { d: 'M6 4v6a6 6 0 0 0 12 0V4' } },
			{ tag: 'line', attrs: { x1: '4', x2: '20', y1: '20', y2: '20' } },
		],
	},
	'home.font.strikethrough': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{ tag: 'path', attrs: { d: 'M16 4H9a3 3 0 0 0-2.83 4' } },
			{ tag: 'path', attrs: { d: 'M14 12a4 4 0 0 1 0 8H6' } },
			{ tag: 'line', attrs: { x1: '4', x2: '20', y1: '12', y2: '12' } },
		],
	},
	// Letters are drawn as filled outlines, never <text>: text in an icon would become part of the
	// button's shown label, which its name (Text Shadow, Character Spacing) would then have to
	// contain (WCAG 2.5.3). The outlines are the bold system-font letters the icons used to draw.
	'home.font.shadow': {
		attrs: { viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': '2' },
		children: [
			{ tag: 'path', attrs: { d: SHADOW_S, fill: 'currentColor', stroke: 'none' } },
			{
				tag: 'path',
				attrs: {
					d: SHADOW_S,
					transform: 'translate(1.5 1.5)',
					'stroke-width': '0.5',
					opacity: '0.4',
				},
			},
		],
	},
	'home.font.increaseFontSize': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{ tag: 'path', attrs: { d: 'M3.5 13h6' } },
			{ tag: 'path', attrs: { d: 'm2 16 4.5-9 4.5 9' } },
			{ tag: 'path', attrs: { d: 'M18 16V7' } },
			{ tag: 'path', attrs: { d: 'm14 11 4-4 4 4' } },
		],
	},
	'home.font.decreaseFontSize': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{ tag: 'path', attrs: { d: 'M3.5 13h6' } },
			{ tag: 'path', attrs: { d: 'm2 16 4.5-9 4.5 9' } },
			{ tag: 'path', attrs: { d: 'M18 7v9' } },
			{ tag: 'path', attrs: { d: 'm14 12 4 4 4-4' } },
		],
	},
	'home.font.clearFormatting': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{ tag: 'path', attrs: { d: 'M4 7V4h16v3' } },
			{ tag: 'path', attrs: { d: 'M5 20h6' } },
			{ tag: 'path', attrs: { d: 'M13 4 8 20' } },
			{ tag: 'path', attrs: { d: 'm15 15 5 5' } },
			{ tag: 'path', attrs: { d: 'm20 15-5 5' } },
		],
	},
	'home.font.characterSpacing': {
		attrs: { viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': '1.5' },
		children: [
			{ tag: 'path', attrs: { d: SPACING_AV, fill: 'currentColor', stroke: 'none' } },
			{ tag: 'path', attrs: { d: 'M3 20 L1 20 M3 20 L5 20', 'stroke-width': '1.5' } },
			{ tag: 'path', attrs: { d: 'M21 20 L19 20 M21 20 L23 20', 'stroke-width': '1.5' } },
		],
	},
	'home.font.changeCase': {
		attrs: { viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': '1.5' },
		children: [
			{
				tag: 'text',
				attrs: {
					x: '2',
					y: '16',
					'font-size': '13',
					'font-weight': 'bold',
					fill: 'currentColor',
					stroke: 'none',
				},
				text: 'Aa',
			},
		],
	},
	'home.font.fontColor': {
		attrs: {
			viewBox: '0 0 24 24',
			fill: 'none',
			stroke: 'currentColor',
			'stroke-width': '2',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [{ tag: 'path', attrs: { d: 'M6 20h12M9.5 4h5L18 16H6L9.5 4z' } }],
	},
	'home.font.highlightColor': {
		attrs: {
			stroke: 'currentColor',
			fill: 'none',
			'stroke-width': '2',
			viewBox: '0 0 24 24',
			'stroke-linecap': 'round',
			'stroke-linejoin': 'round',
		},
		children: [
			{ tag: 'path', attrs: { d: 'm9 11-6 6v3h9l3-3' } },
			{ tag: 'path', attrs: { d: 'm22 12-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4' } },
		],
	},
};
