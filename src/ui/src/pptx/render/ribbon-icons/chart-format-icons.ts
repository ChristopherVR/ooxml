import type { RibbonIconArtwork } from './types';

const STROKE = {
	stroke: 'currentColor',
	fill: 'none',
	'stroke-width': '2',
	viewBox: '0 0 24 24',
	'stroke-linecap': 'round',
	'stroke-linejoin': 'round',
} as const;

/** Chart Format > Current Selection artwork. */
export const CHART_FORMAT_ICONS: Record<string, RibbonIconArtwork> = {
	'chartFormat.currentSelection.formatSelection': {
		attrs: STROKE,
		children: [
			{ tag: 'rect', attrs: { x: '3', y: '3', width: '18', height: '18', rx: '2' } },
			{ tag: 'path', attrs: { d: 'M15 3v18' } },
			{ tag: 'path', attrs: { d: 'M7 15V11M10 15V8' } },
		],
	},
	'chartFormat.currentSelection.resetToMatchStyle': {
		attrs: STROKE,
		children: [
			{ tag: 'path', attrs: { d: 'M3 12a9 9 0 1 0 3-6.7L3 8' } },
			{ tag: 'path', attrs: { d: 'M3 3v5h5' } },
		],
	},
};
