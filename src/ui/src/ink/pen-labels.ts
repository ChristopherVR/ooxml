import type { InkPen } from 'ooxml-core/drawingml';

/** Office's ink colours, offered for a pen. */
export const INK_PEN_COLORS: readonly { value: string; name: string }[] = [
	{ value: '#000000', name: 'Black' },
	{ value: '#FFFFFF', name: 'White' },
	{ value: '#808080', name: 'Gray' },
	{ value: '#E81123', name: 'Red' },
	{ value: '#FF8C00', name: 'Orange' },
	{ value: '#FFC114', name: 'Yellow' },
	{ value: '#FFFF00', name: 'Highlight Yellow' },
	{ value: '#16C60C', name: 'Green' },
	{ value: '#00FFFF', name: 'Turquoise' },
	{ value: '#0078D4', name: 'Blue' },
	{ value: '#8764B8', name: 'Purple' },
	{ value: '#FF00FF', name: 'Pink' },
];

/** The strip's texts, translated by the product. `{kind}`, `{size}` and `{color}` fill a pen name. */
export interface InkPensLabels {
	group: string;
	/** "{kind}: {size} mm, {color}" */
	pen: string;
	kinds: Record<InkPen['kind'], string>;
	thickness: string;
	colors: string;
	remove: string;
	moreColors: string;
	/** Colour names by English name (`Black`, `Red`...). */
	colorNames?: Readonly<Record<string, string>>;
	/** Tooltip line under a pen's name. */
	description: string;
}

export const DEFAULT_INK_PENS_LABELS: InkPensLabels = {
	group: 'Pens',
	pen: '{kind}: {size} mm, {color}',
	kinds: { pen: 'Pen', pencil: 'Pencil', highlighter: 'Highlighter', actionPen: 'Action Pen' },
	thickness: 'Thickness',
	colors: 'Colors',
	remove: 'Delete Pen',
	moreColors: 'More Colors...',
	description: 'Draw ink with this pen. Select it again to change its color and thickness.',
};
