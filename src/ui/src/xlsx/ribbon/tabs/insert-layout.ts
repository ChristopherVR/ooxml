// Insert and Page Layout tabs.
import type { RibbonMenuItem, RibbonTab } from 'ooxml-core/xlsx/ui';
import { ALIGNMENTS } from '../../commands/arrange';
import { shapeSections } from '../../commands/insert-shapes';
import { CHART_TYPES } from '../../commands/insert';
import { MARGIN_PRESETS, PAPER_SIZES } from 'ooxml-core/xlsx/ui';
import { themeFontItems, themePaletteItems, themePresetItems } from 'ooxml-core/xlsx/ui';

const sep: RibbonMenuItem = { separator: true };
const fitOptions = [
	{ value: '0', label: 'Automatic' },
	...Array.from({ length: 9 }, (_, index) => ({
		value: String(index + 1),
		label: String(index + 1),
	})),
];

export function insertTab(): RibbonTab {
	return {
		id: 'insert',
		label: 'Insert',
		groups: [
			{
				id: 'tables',
				label: 'Tables',
				controls: [
					{ kind: 'button', command: 'insert.table', size: 'large' },
					{ kind: 'button', command: 'insert.pivot-table', size: 'large' },
					{ kind: 'button', command: 'insert.recommended-pivot-tables', size: 'large' },
					{
						kind: 'menu',
						label: 'PivotTable',
						items: [
							{ command: 'pivot.settings' },
							{ command: 'pivot.filter-field' },
							{ command: 'pivot.refresh' },
							{ command: 'pivot.group' },
							{ command: 'pivot.group-selection' },
							{ command: 'pivot.ungroup' },
							{ command: 'pivot.field-settings' },
							{ command: 'pivot.delete' },
						],
					},
				],
			},
			{
				id: 'illustrations',
				label: 'Illustrations',
				controls: [
					{ kind: 'button', command: 'insert.pictures', size: 'large' },
					{ kind: 'gallery', command: 'insert.shapes', items: () => [], sections: shapeSections },
					{
						kind: 'menu',
						label: 'Screenshot',
						items: [{ command: 'insert.screenshot' }, { command: 'insert.screen-clipping' }],
					},
					{ kind: 'button', command: 'insert.icons', size: 'large' },
					{ kind: 'button', command: 'insert.3d-models', size: 'large' },
					{ kind: 'button', command: 'insert.smartart', size: 'large' },
				],
			},
			{
				id: 'charts',
				label: 'Charts',
				launcher: 'insert.chart',
				controls: [
					{ kind: 'button', command: 'insert.chart', size: 'large' },
					{ kind: 'button', command: 'insert.pivot-chart', size: 'large' },
					{
						kind: 'stack',
						controls: CHART_TYPES.slice(0, 3).map(([type]) => ({
							kind: 'button' as const,
							command: `insert.chart-${type}`,
						})),
					},
					{
						kind: 'stack',
						controls: CHART_TYPES.slice(3, 6).map(([type]) => ({
							kind: 'button' as const,
							command: `insert.chart-${type}`,
						})),
					},
				],
			},
			{
				id: 'filters',
				label: 'Filters',
				controls: [
					{ kind: 'button', command: 'insert.slicer', size: 'large' },
					{ kind: 'button', command: 'insert.timeline', size: 'large' },
				],
			},
			{
				id: 'links',
				label: 'Links',
				controls: [{ kind: 'button', command: 'insert.link', size: 'large' }],
			},
			{
				id: 'forms',
				label: 'Forms',
				controls: [{ kind: 'button', command: 'insert.forms', size: 'large' }],
			},
			{
				id: 'add-ins',
				label: 'Add-ins',
				controls: [
					{ kind: 'button', command: 'insert.get-add-ins', size: 'large' },
					{ kind: 'button', command: 'insert.my-add-ins', size: 'large' },
				],
			},
			{
				id: 'sparklines',
				label: 'Sparklines',
				launcher: 'sparkline.settings',
				controls: [
					{ kind: 'button', command: 'insert.sparkline-line', size: 'large' },
					{ kind: 'button', command: 'insert.sparkline-column', size: 'large' },
					{ kind: 'button', command: 'insert.sparkline-stacked', size: 'large' },
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'sparkline.settings' },
							{ kind: 'button', command: 'sparkline.group' },
							{ kind: 'button', command: 'sparkline.ungroup' },
							{
								kind: 'menu',
								label: 'Clear Sparklines',
								items: [
									{ command: 'sparkline.clear-selected' },
									{ command: 'sparkline.clear-groups' },
								],
							},
						],
					},
				],
			},
			{
				id: 'comments',
				label: 'Comments',
				controls: [
					{ kind: 'button', command: 'insert.comment', size: 'large' },
					{ kind: 'button', command: 'insert.note', size: 'large' },
				],
			},
			{
				id: 'text',
				label: 'Text',
				controls: [
					{ kind: 'button', command: 'insert.text-box', size: 'large' },
					{ kind: 'button', command: 'insert.header-footer', size: 'large' },
				],
			},
			{
				id: 'symbols',
				label: 'Symbols',
				controls: [
					{ kind: 'button', command: 'insert.word-art', size: 'large' },
					{ kind: 'button', command: 'insert.object', size: 'large' },
					{ kind: 'button', command: 'insert.symbol', size: 'large' },
					{ kind: 'button', command: 'insert.equation', size: 'large' },
					{ kind: 'button', command: 'drawing.edit-equation' },
				],
			},
		],
	};
}

export function pageLayoutTab(): RibbonTab {
	return {
		id: 'page-layout',
		label: 'Page Layout',
		groups: [
			{
				id: 'themes',
				label: 'Themes',
				controls: [
					{ kind: 'gallery', command: 'page.themes', items: themePresetItems },
					{ kind: 'gallery', command: 'page.theme-colors', items: themePaletteItems },
					{ kind: 'gallery', command: 'page.theme-fonts', items: themeFontItems },
				],
			},
			{
				id: 'page-setup',
				label: 'Page Setup',
				launcher: 'page.setup',
				controls: [
					{
						kind: 'menu',
						label: 'Margins',
						icon: 'xl-margins',
						size: 'large',
						items: [
							...MARGIN_PRESETS.map(([id, label]) => ({ command: 'page.margins', arg: id, label })),
							sep,
							{ command: 'page.margins-custom' },
						],
					},
					{
						kind: 'menu',
						label: 'Orientation',
						icon: 'xl-pageOrientation',
						size: 'large',
						items: [
							{ command: 'page.orientation-portrait' },
							{ command: 'page.orientation-landscape' },
						],
					},
					{
						kind: 'menu',
						label: 'Size',
						icon: 'xl-pageSize',
						size: 'large',
						items: [
							...PAPER_SIZES.map(([code, label]) => ({ command: 'page.size', arg: code, label })),
							sep,
							{ command: 'page.size-more' },
						],
					},
					{
						kind: 'menu',
						label: 'Print Area',
						icon: 'xl-printArea',
						size: 'large',
						items: [{ command: 'page.print-area-set' }, { command: 'page.print-area-clear' }],
					},
					{
						kind: 'menu',
						label: 'Breaks',
						icon: 'xl-printArea',
						size: 'large',
						items: [
							{ command: 'page.break-row-insert' },
							{ command: 'page.break-col-insert' },
							sep,
							{ command: 'page.break-row-remove' },
							{ command: 'page.break-col-remove' },
							sep,
							{ command: 'page.breaks-reset' },
						],
					},
					{ kind: 'button', command: 'page.preview', size: 'large' },
					{ kind: 'button', command: 'page.print-titles', size: 'large' },
					{ kind: 'button', command: 'page.background', size: 'large' },
					{ kind: 'button', command: 'page.background-delete', size: 'large' },
				],
			},
			{
				id: 'scale-to-fit',
				label: 'Scale to Fit',
				launcher: 'page.setup',
				controls: [
					{
						kind: 'stack',
						controls: [
							{ kind: 'select', command: 'page.fit-width', options: fitOptions, showLabel: true },
							{ kind: 'select', command: 'page.fit-height', options: fitOptions, showLabel: true },
							{
								kind: 'select',
								command: 'page.scale',
								options: [],
								editable: true,
								showLabel: true,
								width: 58,
							},
						],
					},
				],
			},
			{
				id: 'arrange',
				label: 'Arrange',
				controls: [
					{ kind: 'button', command: 'arrange.forward', size: 'large' },
					{ kind: 'button', command: 'arrange.backward', size: 'large' },
					{ kind: 'button', command: 'arrange.selection-pane', size: 'large' },
					{
						kind: 'stack',
						controls: [
							{
								kind: 'menu',
								label: 'Align',
								items: ALIGNMENTS.map(([id]) => ({ command: 'arrange.align-' + id })),
							},
							{
								kind: 'menu',
								label: 'Group',
								items: [{ command: 'arrange.group' }, { command: 'arrange.ungroup' }],
							},
							{
								kind: 'menu',
								label: 'Rotate',
								items: [
									{ command: 'arrange.rotate-right' },
									{ command: 'arrange.rotate-left' },
									{ command: 'arrange.flip-horizontal' },
									{ command: 'arrange.flip-vertical' },
								],
							},
						],
					},
				],
			},
			{
				id: 'sheet-options',
				label: 'Sheet Options',
				launcher: 'page.setup',
				controls: [
					{
						kind: 'stack',
						controls: [
							{ kind: 'toggle', command: 'page.gridlines-view', showLabel: true },
							{ kind: 'toggle', command: 'page.gridlines-print', showLabel: true },
						],
					},
					{
						kind: 'stack',
						controls: [
							{ kind: 'toggle', command: 'page.headings-view', showLabel: true },
							{ kind: 'toggle', command: 'page.headings-print', showLabel: true },
						],
					},
				],
			},
		],
	};
}
