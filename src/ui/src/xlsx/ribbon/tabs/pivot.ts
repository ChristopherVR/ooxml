import { DEFAULT_THEME } from 'ooxml-core/xlsx';
import {
	tableStylePreview,
	type EditorContext,
	type RibbonControl,
	type RibbonTab,
} from 'ooxml-core/xlsx/ui';
import { selectedPivot } from '../../commands/pivots';

const contextual = (ctx: EditorContext) => !!selectedPivot(ctx);
const button = (command: string): RibbonControl => ({ kind: 'button', command, showLabel: true });
const toggle = (command: string): RibbonControl => ({ kind: 'toggle', command, showLabel: true });
const stack = (...controls: RibbonControl[]): RibbonControl => ({ kind: 'stack', controls });
const menu = (label: string, ...ids: string[]): RibbonControl => ({
	kind: 'menu',
	label,
	icon: 'xl-table',
	size: 'large',
	items: ids.map((command) => ({ command: `pivot.${command}` })),
});

export function pivotAnalyzeTab(): RibbonTab {
	return {
		id: 'pivot-analyze',
		label: 'PivotTable Analyze',
		contextual,
		deferControls: true,
		groups: [
			{
				id: 'pivot-table',
				label: 'PivotTable',
				controls: [stack(button('pivot.name'), button('pivot.settings'))],
			},
			{
				id: 'active-field',
				label: 'Active Field',
				controls: [
					stack(
						button('pivot.field-settings'),
						button('pivot.expand-field'),
						button('pivot.collapse-field'),
					),
					stack(button('pivot.drill-down'), button('pivot.drill-up')),
				],
			},
			{
				id: 'group',
				label: 'Group',
				controls: [
					stack(button('pivot.group-selection'), button('pivot.ungroup'), button('pivot.group')),
				],
			},
			{
				id: 'filter',
				label: 'Filter',
				controls: [
					stack(button('insert.slicer'), button('insert.timeline'), button('pivot.connections')),
					stack(button('pivot.condition'), button('pivot.filter-field')),
				],
			},
			{
				id: 'data',
				label: 'Data',
				controls: [
					{
						kind: 'split',
						command: 'pivot.refresh',
						label: 'Refresh',
						size: 'large',
						menu: [
							{ command: 'pivot.refresh', label: 'Refresh' },
							{ command: 'pivot.refresh-all', label: 'Refresh All' },
						],
					},
					button('pivot.source'),
				],
			},
			{
				id: 'actions',
				label: 'Actions',
				controls: [
					menu('Clear', 'clear-filters', 'delete'),
					menu('Select', 'select'),
					button('pivot.move'),
				],
			},
			{
				id: 'calculations',
				label: 'Calculations',
				controls: [menu('Fields, Items & Sets', 'calculated-field')],
			},
			{
				id: 'tools',
				label: 'Tools',
				controls: [stack(button('insert.pivot-chart'), button('insert.recommended-pivot-tables'))],
			},
			{
				id: 'show',
				label: 'Show',
				controls: [
					stack(
						button('pivot.field-list'),
						toggle('pivot.show-drill'),
						toggle('pivot.show-headers'),
					),
				],
			},
		],
	};
}

const styleGallery = (ctx: EditorContext) => {
	if (!selectedPivot(ctx)) return [];
	const theme = ctx.workbook()?.theme ?? DEFAULT_THEME;
	return [
		{ id: '', label: 'None' },
		...(['Light', 'Medium', 'Dark'] as const).flatMap((family) =>
			Array.from({ length: 28 }, (_, at) => {
				const index = at + 1;
				const name = `PivotStyle${family}${index}`;
				return {
					id: name,
					label: ctx.t('{family} Style {n}', { family: ctx.t(family), n: index }),
					preview: tableStylePreview({ name, family, index }, theme),
				};
			}),
		),
	];
};

export function pivotDesignTab(): RibbonTab {
	return {
		id: 'pivot-design',
		label: 'PivotTable Design',
		contextual,
		deferControls: true,
		groups: [
			{
				id: 'layout',
				label: 'Layout',
				controls: [
					menu('Subtotals', 'subtotal-none', 'subtotal-bottom', 'subtotal-top'),
					menu('Grand Totals', 'totals-off', 'totals-on', 'totals-rows', 'totals-columns'),
					menu(
						'Report Layout',
						'layout-compact',
						'layout-outline',
						'layout-tabular',
						'repeat-labels',
						'no-repeat-labels',
					),
					menu('Blank Rows', 'insert-blank-rows', 'remove-blank-rows'),
				],
			},
			{
				id: 'pivot-style-options',
				label: 'PivotTable Style Options',
				controls: [
					stack(toggle('pivot.showRowHeaders'), toggle('pivot.showColHeaders')),
					stack(toggle('pivot.showRowStripes'), toggle('pivot.showColStripes')),
				],
			},
			{
				id: 'pivot-styles',
				label: 'PivotTable Styles',
				controls: [{ kind: 'gallery', command: 'pivot.style', items: styleGallery }],
			},
		],
	};
}
