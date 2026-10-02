// Insert and Page Layout tabs.
import type { RibbonMenuItem, RibbonTab } from '../parts.js';
import { CHART_TYPES } from '../../commands/insert.js';
import { MARGIN_PRESETS, PAPER_SIZES } from '../../commands/page-layout.js';
import { themePaletteItems } from './theme-palette.js';

const sep: RibbonMenuItem = { separator: true };

export function insertTab(): RibbonTab {
	return {
		id: 'insert',
		label: 'Insert',
		groups: [
			{
				id: 'tables',
				label: 'Tables',
				controls: [{ kind: 'button', command: 'insert.table', size: 'large' }],
			},
			{
				id: 'illustrations',
				label: 'Illustrations',
				controls: [{ kind: 'button', command: 'insert.pictures', size: 'large' }],
			},
			{
				id: 'charts',
				label: 'Charts',
				launcher: 'insert.chart',
				controls: [
					{ kind: 'button', command: 'insert.chart', size: 'large' },
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
				id: 'links',
				label: 'Links',
				controls: [{ kind: 'button', command: 'insert.link', size: 'large' }],
			},
			{
				id: 'comments',
				label: 'Comments',
				controls: [{ kind: 'button', command: 'insert.comment', size: 'large' }],
			},
			{
				id: 'symbols',
				label: 'Symbols',
				controls: [{ kind: 'button', command: 'insert.symbol', size: 'large' }],
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
				controls: [{ kind: 'gallery', command: 'page.theme-colors', items: themePaletteItems }],
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
