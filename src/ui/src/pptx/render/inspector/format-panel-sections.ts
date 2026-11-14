/**
 * Which Format pane section holds each existing inspector panel, so the five
 * bindings place their panels the same way.
 *
 * @module render/inspector/format-panel-sections
 */
import type { FormatSectionId } from './format-sections';

/**
 * Where each binding's existing panels go. A binding whose panel combines two
 * concerns splits it along these lines (React's fill-and-stroke card renders
 * its fill part under `fill` and its line part under `outline`).
 */
export const FORMAT_PANEL_SECTIONS = {
	slideProperties: 'slide',
	slideBackground: 'slide',
	text: 'font',
	textEffects: 'font',
	textBox: 'paragraph',
	image: 'picture',
	table: 'table',
	chart: 'chart',
	smartArt: 'smartArt',
	media: 'media',
	ole: 'object',
	quickStyles: 'fill',
	fill: 'fill',
	stroke: 'outline',
	connectorArrows: 'outline',
	effects: 'effects',
	shapeType: 'arrange',
	groupInfo: 'arrange',
	layerOrder: 'arrange',
	lock: 'arrange',
	geometry: 'sizePosition',
	transform: 'sizePosition',
	accessibility: 'altText',
	actionSettings: 'action',
} as const satisfies Record<string, FormatSectionId>;

export type FormatPanelId = keyof typeof FORMAT_PANEL_SECTIONS;

/** Whether the binding's `panel` belongs in `section`. */
export function formatPanelIn(panel: FormatPanelId, section: FormatSectionId): boolean {
	return FORMAT_PANEL_SECTIONS[panel] === section;
}
