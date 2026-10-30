import type { VisioDataGraphicItem, VisioDataGraphicKind } from 'ooxml-core/visio/ui';

/** Display type of a data graphic by the command (or dialog value) that makes it. */
export const GRAPHIC_KINDS: Record<string, VisioDataGraphicKind> = {
	'graphic-text': 'text',
	'graphic-bar': 'bar',
	'graphic-icon': 'icon',
	'graphic-color': 'color',
};
export const KIND_LABELS = ['Text Callout', 'Data Bar', 'Icon Set', 'Color by Value'];
export const POSITION_LABELS = [
	'Top left',
	'Top center',
	'Top right',
	'Left',
	'Right',
	'Bottom left',
	'Bottom center',
	'Bottom right',
];
export const kindLabel = (kind: VisioDataGraphicKind): string =>
	KIND_LABELS[Object.values(GRAPHIC_KINDS).indexOf(kind)] ?? kind;

/** One line of the Edit Data Graphic list: the display type, the field and where it sits. */
export const itemLabel = (item: VisioDataGraphicItem, fieldLabel: string): string =>
	`${kindLabel(item.kind)}: ${fieldLabel}${item.position ? ` (${item.position.replace('-', ' ')})` : ''}`;
