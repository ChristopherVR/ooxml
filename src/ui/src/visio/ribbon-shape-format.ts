import { arrangeGroup, shapeStylesGroup } from './ribbon-arrange';
import { fontColorMenu } from './ribbon-home-format';
import { command, commandRow, group, menu, stack } from './ribbon-parts';
import { sizeFields } from './ribbon-size-fields';

import { colorGrid } from './ribbon-color-menu';

/** Contextual shape styles, label formatting, Arrange, and Size & Position. */
export function buildShapeFormatPanel(doc: Document, panel: HTMLElement): void {
	const outline = menu(doc, {
		id: 'text-outline',
		label: 'Text Outline',
		icon: 'line',
		size: 'small',
		split: true,
		action: { type: 'text-outline' },
		items: [],
	});
	outline.append(colorGrid(doc, 'text-outline'));
	const effects = menu(doc, {
		id: 'text-effects',
		label: 'Text Effects',
		icon: 'effects',
		size: 'small',
		items: [
			{
				id: 'text-shadow',
				label: 'Shadow',
				action: { type: 'text-effect', effect: 'textShadow', enabled: true },
			},
			{
				id: 'no-text-shadow',
				label: 'No Shadow',
				action: { type: 'text-effect', effect: 'textShadow', enabled: false },
			},
			{
				id: 'text-glow',
				label: 'Glow',
				action: { type: 'text-effect', effect: 'textGlow', enabled: true },
			},
			{
				id: 'no-text-glow',
				label: 'No Glow',
				action: { type: 'text-effect', effect: 'textGlow', enabled: false },
			},
		],
	});
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Shape Format commands', [
			shapeStylesGroup(doc),
			group(doc, 'WordArt Styles', [
				stack(doc, [fontColorMenu(doc, 'Text Fill'), outline, effects]),
			]),
			arrangeGroup(doc),
			group(doc, 'Size', [
				sizeFields(doc),
				command(doc, {
					id: 'size-position',
					label: 'Size & Position',
					icon: 'sizePosition',
					action: { type: 'sizePosition' },
					checked: false,
				}),
			]),
		]),
	);
}
