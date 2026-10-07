// Contextual tabs: Table Design (active cell inside a table) and Chart Design (active cell under
// a chart).
import type { RibbonTab } from 'ooxml-core/xlsx/ui';
import { activeChart, activeTable } from 'ooxml-core/xlsx/ui';
import { CHART_TYPES } from '../../commands/insert';
import { tableStyleGalleryItems } from '../../commands/styles';
import { chartColorGalleryItems, chartColorGallerySections } from '../../commands/chart-colors';

export function tableDesignTab(): RibbonTab {
	return {
		id: 'table-design',
		label: 'Table Design',
		contextual: (ctx) => !!activeTable(ctx),
		groups: [
			{
				id: 'properties',
				label: 'Properties',
				controls: [
					{ kind: 'button', command: 'table.name', size: 'small', showLabel: true },
					{ kind: 'button', command: 'table.resize', size: 'small' },
				],
			},
			{
				id: 'tools',
				label: 'Tools',
				controls: [{ kind: 'button', command: 'table.convert-to-range', size: 'large' }],
			},
			{
				id: 'style-options',
				label: 'Table Style Options',
				controls: [
					{
						kind: 'stack',
						controls: [
							{ kind: 'toggle', command: 'table.header-row' },
							{ kind: 'toggle', command: 'table.total-row' },
							{ kind: 'toggle', command: 'table.banded-rows' },
						],
					},
					{
						kind: 'stack',
						controls: [
							{ kind: 'toggle', command: 'table.first-column' },
							{ kind: 'toggle', command: 'table.last-column' },
							{ kind: 'toggle', command: 'table.banded-columns' },
						],
					},
				],
			},
			{
				id: 'table-styles',
				label: 'Table Styles',
				controls: [{ kind: 'gallery', command: 'table.style', items: tableStyleGalleryItems }],
			},
		],
	};
}

export function chartDesignTab(): RibbonTab {
	return {
		id: 'chart-design',
		label: 'Chart Design',
		contextual: (ctx) => !!activeChart(ctx),
		groups: [
			{
				id: 'chart-styles',
				label: 'Chart Styles',
				controls: [
					{
						kind: 'gallery',
						command: 'chart.colors',
						items: chartColorGalleryItems,
						sections: chartColorGallerySections,
					},
				],
			},
			{
				id: 'chart-layouts',
				label: 'Chart Layouts',
				controls: [
					{ kind: 'button', command: 'chart.title', size: 'large' },
					{ kind: 'toggle', command: 'chart.legend', size: 'large' },
				],
			},
			{
				id: 'type',
				label: 'Type',
				controls: [
					{
						kind: 'split',
						command: 'chart.type',
						size: 'large',
						menu: CHART_TYPES.map(([type, label]) => ({ command: 'chart.type', arg: type, label })),
					},
				],
			},
			{
				id: 'series-options',
				label: 'Series Options',
				controls: [{ kind: 'button', command: 'chart.format-series', size: 'large' }],
			},
			{
				id: 'chart-delete',
				label: 'Delete',
				controls: [{ kind: 'button', command: 'chart.delete', size: 'large' }],
			},
		],
	};
}
