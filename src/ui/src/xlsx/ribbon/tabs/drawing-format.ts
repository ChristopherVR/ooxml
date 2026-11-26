import type { RibbonControl, RibbonTab } from 'ooxml-core/xlsx/ui';
import { formatTabContext, type DrawingFormatAction } from '../../commands/drawing-format';
import { shapeSections } from '../../commands/insert-shapes';
import { arrangeGroup } from './arrange';

const button = (command: string): RibbonControl => ({ kind: 'button', command, size: 'large' });
const shapes = (command: string): RibbonControl => ({
	kind: 'gallery',
	command,
	items: () => [],
	sections: shapeSections,
});
const effects = (label: string, command: DrawingFormatAction): RibbonControl => ({
	kind: 'menu',
	label,
	size: 'large',
	items: [
		{ command, label: 'Drop Shadow', arg: true },
		{ command, label: 'No Shadow', arg: false },
	],
});

export function pictureFormatTab(): RibbonTab {
	return {
		id: 'picture-format',
		label: 'Picture Format',
		deferControls: true,
		contextual: (ctx) => formatTabContext(ctx, 'picture'),
		groups: [
			{
				id: 'adjust',
				label: 'Adjust',
				controls: [
					button('picture.corrections'),
					{
						kind: 'menu',
						label: 'Color',
						size: 'large',
						items: [
							{ command: 'picture.color', label: 'Grayscale', arg: true },
							{ command: 'picture.color', label: 'Original Color', arg: false },
						],
					},
					button('picture.artistic-effects'),
					button('picture.compress'),
					button('picture.change'),
					button('picture.reset'),
				],
			},
			{
				id: 'picture-styles',
				label: 'Picture Styles',
				controls: [
					{
						kind: 'gallery',
						command: 'picture.style',
						items: () => [],
						loadItems: async () =>
							(await import('../../commands/drawing-format-styles')).pictureStyleTiles,
					},
					button('picture.border'),
					effects('Picture Effects', 'picture.effects'),
				],
			},
			{ id: 'accessibility', label: 'Accessibility', controls: [button('drawing.alt-text')] },
			arrangeGroup(),
			{ id: 'size', label: 'Size', controls: [button('picture.crop'), button('drawing.size')] },
		],
	};
}

export function shapeFormatTab(): RibbonTab {
	return {
		id: 'shape-format',
		label: 'Shape Format',
		deferControls: true,
		contextual: (ctx) => formatTabContext(ctx, 'shape'),
		groups: [
			{
				id: 'insert-shapes',
				label: 'Insert Shapes',
				controls: [
					shapes('insert.shapes'),
					shapes('drawing.edit-shape'),
					button('insert.text-box'),
				],
			},
			{
				id: 'shape-styles',
				label: 'Shape Styles',
				controls: [
					{
						kind: 'gallery',
						command: 'drawing.shape-style',
						items: () => [],
						loadItems: async (ctx) =>
							(await import('../../commands/drawing-format-styles')).shapeStyleTiles(
								ctx.workbook(),
								ctx.t,
							),
					},
					button('drawing.shape-fill'),
					button('drawing.shape-outline'),
					effects('Shape Effects', 'drawing.shape-effects'),
				],
			},
			{
				id: 'word-art-styles',
				label: 'WordArt Styles',
				controls: [
					{
						kind: 'gallery',
						command: 'drawing.word-art-style',
						items: () => [],
						loadItems: async (ctx) =>
							(await import('../../commands/drawing-format-styles')).wordArtStyleTiles(ctx.t),
					},
					button('drawing.text-fill'),
					button('drawing.text-outline'),
					effects('Text Effects', 'drawing.text-effects'),
				],
			},
			{ id: 'accessibility', label: 'Accessibility', controls: [button('drawing.alt-text')] },
			arrangeGroup(),
			{ id: 'size', label: 'Size', controls: [button('drawing.size')] },
		],
	};
}
