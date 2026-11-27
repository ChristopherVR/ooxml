// Format (chart): Excel's chart Format contextual tab, shown while a chart is the selected object.
// Insert Shapes and Arrange are the Shape Format tab's controls; the style galleries load their
// tiles on first open, as Shape Format's do.
import type { RibbonControl, RibbonTab } from 'ooxml-core/xlsx/ui';
import { activeChart } from 'ooxml-core/xlsx/ui';
import { chartElementOptions } from '../../commands/chart-format-tab';
import { shapeSections } from '../../commands/insert-shapes';
import { arrangeGroup } from './arrange';

const button = (command: string, size: 'large' | 'small' = 'large'): RibbonControl => ({
	kind: 'button',
	command,
	size,
	...(size === 'small' ? { showLabel: true } : {}),
});
const shadow = (label: string, command: string): RibbonControl => ({
	kind: 'menu',
	label,
	items: [
		{ command, label: 'Drop Shadow', arg: true },
		{ command, label: 'No Shadow', arg: false },
	],
});

export function chartFormatTab(): RibbonTab {
	return {
		id: 'chart-format',
		label: 'Format',
		deferControls: true,
		contextual: (ctx) => !!activeChart(ctx),
		groups: [
			{
				id: 'current-selection',
				label: 'Current Selection',
				controls: [
					{
						kind: 'stack',
						controls: [
							{
								kind: 'select',
								command: 'chart.element',
								options: chartElementOptions,
								width: 150,
							},
							button('chart.format-selection', 'small'),
							button('chart.reset-style', 'small'),
						],
					},
				],
			},
			{
				id: 'insert-shapes',
				label: 'Insert Shapes',
				controls: [
					{ kind: 'gallery', command: 'insert.shapes', items: () => [], sections: shapeSections },
					button('insert.text-box'),
				],
			},
			{
				id: 'shape-styles',
				label: 'Shape Styles',
				controls: [
					{
						kind: 'gallery',
						command: 'chart.shape-style',
						columns: 7,
						items: () => [],
						loadItems: async (ctx) =>
							(await import('../../commands/drawing-format-styles')).shapeStyleTiles(
								ctx.workbook(),
								ctx.t,
							),
					},
					{
						kind: 'stack',
						controls: [
							{
								kind: 'color',
								command: 'chart.shape-fill',
								label: 'Shape Fill',
								icon: 'xl-fillColor',
							},
							{
								kind: 'color',
								command: 'chart.shape-outline',
								label: 'Shape Outline',
								icon: 'xl-borders',
							},
							shadow('Shape Effects', 'chart.shape-effects'),
						],
					},
				],
			},
			{
				id: 'word-art-styles',
				label: 'WordArt Styles',
				controls: [
					{
						kind: 'gallery',
						command: 'chart.word-art-style',
						items: () => [],
						loadItems: async (ctx) =>
							(await import('../../commands/drawing-format-styles')).wordArtStyleTiles(ctx.t),
					},
					{
						kind: 'stack',
						controls: [
							{
								kind: 'color',
								command: 'chart.text-fill',
								label: 'Text Fill',
								icon: 'xl-fontColor',
							},
							{
								kind: 'color',
								command: 'chart.text-outline',
								label: 'Text Outline',
								icon: 'xl-fontColor',
							},
						],
					},
				],
			},
			arrangeGroup(),
			{ id: 'size', label: 'Size', controls: [button('chart.size')] },
		],
	};
}
