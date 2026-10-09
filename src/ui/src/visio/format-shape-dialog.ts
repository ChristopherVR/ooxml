import { VISIO_SHADOW_PRESETS } from 'ooxml-core/visio';
import type { PaintButton, PaintDialog, PaintInput } from './viewer-paint-dialog';

/** Format Shape fields: Visio's Fill, Line and Effects sections over the format-shape edit. */
export const FORMAT_SHAPE_FIELDS = {
	Fill: [
		['fillColor', 'Color', 'color'],
		['fillTransparency', 'Transparency (%)', 'percent'],
	],
	Line: [
		['lineColor', 'Color', 'color'],
		['lineWeight', 'Width (pt)', 'points'],
		['linePattern', 'Dash type', 'pattern'],
		['lineTransparency', 'Transparency (%)', 'percent'],
	],
	Effects: [
		['shadow', 'Shadow', 'shadow'],
		['glowSize', 'Glow size (pt)', 'points'],
		['glowColor', 'Glow color', 'color'],
		['glowTransparency', 'Glow transparency (%)', 'percent'],
		['softEdges', 'Soft edges size (pt)', 'points'],
		['reflectionSize', 'Reflection size (%)', 'percent'],
		['reflectionTransparency', 'Reflection transparency (%)', 'percent'],
		['reflectionDistance', 'Reflection distance (pt)', 'points'],
		['reflectionBlur', 'Reflection blur (pt)', 'points'],
	],
} as const;
type Section = keyof typeof FORMAT_SHAPE_FIELDS;
export type FormatShapeField = (typeof FORMAT_SHAPE_FIELDS)[Section][number][0];
export type FormatShapeKind = (typeof FORMAT_SHAPE_FIELDS)[Section][number][2];
export const FORMAT_SHAPE_UNSUPPORTED =
	'Arrows, gradients, bevel and 3-D rotation are not available in this pane: arrowheads have no edit command yet, and bevel and 3-D rotation need a 3-D renderer.';

type Select = Exclude<PaintInput, HTMLInputElement>;

/** The Format Shape dialog, built from shared `office-ui-*` controls. */
export function createFormatShapeDialog(doc: Document) {
	const dialog = doc.createElement('office-ui-dialog') as PaintDialog;
	dialog.className = 'format-shape-dialog paint-properties-dialog';
	dialog.setAttribute('heading', 'Format Shape');
	const fields = new Map<FormatShapeField, PaintInput>();
	const kinds = new Map<FormatShapeField, FormatShapeKind>();
	for (const [heading, entries] of Object.entries(FORMAT_SHAPE_FIELDS)) {
		const group = doc.createElement('fieldset');
		const legend = doc.createElement('legend');
		legend.textContent = heading;
		group.append(legend);
		for (const [field, text, kind] of entries) {
			const label = doc.createElement('label');
			label.textContent = text;
			let input: PaintInput;
			if (kind === 'pattern' || kind === 'shadow') {
				const select = doc.createElement('office-ui-select') as Select;
				select.options =
					kind === 'shadow'
						? VISIO_SHADOW_PRESETS.map((preset) => ({
								value: preset,
								label: preset === 'none' ? 'No Shadow' : `Offset: ${preset.replace('-', ' ')}`,
							}))
						: Array.from({ length: 24 }, (_, index) => ({
								value: String(index),
								label: index === 0 ? 'No Line' : index === 1 ? 'Solid' : `Dash ${index}`,
							}));
				input = select;
			} else {
				const element = doc.createElement('input');
				element.type = kind === 'color' ? 'text' : 'number';
				if (kind === 'color') {
					element.placeholder = '#RRGGBB';
					element.maxLength = 7;
				} else {
					element.min = '0';
					element.max = kind === 'points' ? '150' : '100';
					element.step = '0.5';
				}
				input = element;
			}
			input.dataset.formatField = field;
			input.setAttribute('aria-label', `${heading} ${text}`);
			label.append(input);
			group.append(label);
			fields.set(field, input);
			kinds.set(field, kind);
		}
		dialog.append(group);
	}
	const hint = doc.createElement('p');
	hint.className = 'paint-properties-hint';
	hint.textContent = `Only changed fields are applied; a fill colour of "none" removes the fill. ${FORMAT_SHAPE_UNSUPPORTED}`;
	const error = doc.createElement('p');
	error.setAttribute('role', 'alert');
	error.dataset.formatError = '';
	const buttons = ['Apply', 'Cancel'].map((label) => {
		const button = doc.createElement('office-ui-button') as PaintButton;
		button.slot = 'footer';
		button.setAttribute('label', label);
		button.setAttribute('command', `format-shape-${label.toLowerCase()}`);
		return button;
	});
	dialog.append(hint, error, ...buttons);
	return { dialog, fields, kinds, error, apply: buttons[0]!, cancel: buttons[1]! };
}
