// Contextual tabs: Table Design (active cell inside a table) and Chart Design (active cell under
// a chart).
import type { RibbonTab } from 'ooxml-core/xlsx/ui';
import {
	activeChart,
	activeSlicer,
	activeSmartArt,
	activeTable,
	activeTimeline,
} from 'ooxml-core/xlsx/ui';
import {
	DIAGRAM_ADD_SHAPE_COMMANDS,
	DIAGRAM_ADD_SHAPE_LABEL,
	DIAGRAM_DESIGN_GROUPS,
} from 'ooxml-core/diagram';
import { smartArtCommandId } from '../../smartart-controller';
import { smartArtColorsGalleryState, smartArtStylesGalleryState } from '../../../smartart';
import type { DiagramColorScheme, DiagramStyleIntensity } from 'ooxml-core/diagram';
import { smartArtSchemeColors } from '../../grid/smartart';
import { gallerySections } from '../shared-gallery';
import { SLICER_SIZE_OPTIONS, slicerStyleGalleryItems } from '../../commands/slicers';
import { CHART_TYPES } from '../../commands/insert';
import { tableStyleGalleryItems } from '../../commands/styles';
import {
	chartColorGalleryItems,
	chartColorGallerySections,
	chartQuickLayoutGalleryItems,
	chartQuickLayoutGallerySections,
	chartStyleGalleryItems,
	chartStyleGallerySections,
} from '../../commands/chart-design';

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

const SLICER_COLUMN_OPTIONS = Array.from({ length: 10 }, (_v, i) => ({
	value: String(i + 1),
	label: String(i + 1),
}));
const SLICER_SIZE_CHOICES = SLICER_SIZE_OPTIONS.map((value) => ({ value, label: value }));

/** Slicer: shown while a slicer is the selected object, as in Excel's Slicer tab. */
export function slicerTab(): RibbonTab {
	return {
		id: 'slicer',
		label: 'Slicer',
		contextual: (ctx) => !!activeSlicer(ctx),
		groups: [
			{
				id: 'slicer',
				label: 'Slicer',
				controls: [{ kind: 'button', command: 'slicer.settings', size: 'large' }],
			},
			{
				id: 'slicer-styles',
				label: 'Slicer Styles',
				controls: [{ kind: 'gallery', command: 'slicer.style', items: slicerStyleGalleryItems }],
			},
			{
				id: 'buttons',
				label: 'Buttons',
				controls: [
					{
						kind: 'stack',
						controls: [
							{
								kind: 'select',
								command: 'slicer.columns',
								options: SLICER_COLUMN_OPTIONS,
								width: 56,
								editable: true,
							},
							{
								kind: 'select',
								command: 'slicer.button-height',
								options: SLICER_SIZE_CHOICES,
								width: 64,
								editable: true,
							},
						],
					},
				],
			},
			{
				id: 'size',
				label: 'Size',
				controls: [
					{
						kind: 'stack',
						controls: [
							{
								kind: 'select',
								command: 'slicer.height',
								options: SLICER_SIZE_CHOICES,
								width: 64,
								editable: true,
							},
							{
								kind: 'select',
								command: 'slicer.width',
								options: SLICER_SIZE_CHOICES,
								width: 64,
								editable: true,
							},
						],
					},
				],
			},
			{
				id: 'slicer-delete',
				label: 'Delete',
				controls: [{ kind: 'button', command: 'slicer.delete', size: 'large' }],
			},
		],
	};
}

const TIMELINE_LEVEL_OPTIONS = [
	{ value: '0', label: 'Years' },
	{ value: '1', label: 'Quarters' },
	{ value: '2', label: 'Months' },
	{ value: '3', label: 'Days' },
];

/** Timeline: shown while a timeline is the selected object. */
export function timelineTab(): RibbonTab {
	return {
		id: 'timeline',
		label: 'Timeline',
		contextual: (ctx) => !!activeTimeline(ctx),
		groups: [
			{
				id: 'timeline',
				label: 'Timeline',
				controls: [{ kind: 'button', command: 'timeline.settings', size: 'large' }],
			},
			{
				id: 'time-level',
				label: 'Show',
				controls: [
					{
						kind: 'select',
						command: 'timeline.level',
						options: TIMELINE_LEVEL_OPTIONS,
						width: 96,
					},
				],
			},
			{
				id: 'timeline-delete',
				label: 'Delete',
				controls: [{ kind: 'button', command: 'timeline.delete', size: 'large' }],
			},
		],
	};
}

/**
 * SmartArt Design: shown while a SmartArt graphic is the selected object, with the same commands
 * as Word's tab (the shared list of `ooxml-core/diagram`). SmartArt Styles and Change Colors are
 * not offered: only one colour and one style definition ship with the core.
 */
export function smartArtDesignTab(): RibbonTab {
	const small = (key: Parameters<typeof smartArtCommandId>[0]) => ({
		kind: 'button' as const,
		command: smartArtCommandId(key),
		showLabel: true,
	});
	return {
		id: 'smartart-design',
		label: 'SmartArt Design',
		contextual: (ctx) => !!activeSmartArt(ctx),
		groups: [
			{
				id: 'create-graphic',
				label: DIAGRAM_DESIGN_GROUPS.create,
				controls: [
					{
						kind: 'menu',
						label: DIAGRAM_ADD_SHAPE_LABEL,
						icon: 'xl-smartArtAddShape',
						size: 'large',
						items: DIAGRAM_ADD_SHAPE_COMMANDS.map((command) => ({
							command: smartArtCommandId(command.key),
						})),
					},
					{ kind: 'toggle', command: smartArtCommandId('textPane'), size: 'large' },
					{ kind: 'stack', controls: [small('promote'), small('demote')] },
					{ kind: 'stack', controls: [small('moveUp'), small('moveDown')] },
				],
			},
			{
				id: 'layouts',
				label: DIAGRAM_DESIGN_GROUPS.layouts,
				controls: [{ kind: 'button', command: smartArtCommandId('changeLayout'), size: 'large' }],
			},
			{
				id: 'smartart-format',
				label: 'SmartArt Styles',
				controls: [
					{
						kind: 'gallery',
						command: 'smartart.style',
						items: () => [],
						sections: (ctx) =>
							gallerySections(
								smartArtStylesGalleryState(
									smartArtSchemeColors(ctx.workbook()?.theme).accent1,
									activeSmartArt(ctx)?.smartArt.quickStyleId?.split('/').at(-1) as
										| DiagramStyleIntensity
										| undefined,
									(key) => ctx.t(key),
								).sections,
							),
					},
					{
						kind: 'gallery',
						command: 'smartart.colors',
						items: () => [],
						sections: (ctx) =>
							gallerySections(
								smartArtColorsGalleryState(
									smartArtSchemeColors(ctx.workbook()?.theme),
									activeSmartArt(ctx)?.smartArt.colorsId?.split('/').at(-1) as
										| DiagramColorScheme
										| undefined,
									(key) => ctx.t(key),
								).sections,
							),
					},
				],
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
						command: 'chart.styles',
						items: chartStyleGalleryItems,
						sections: chartStyleGallerySections,
					},
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
					{
						kind: 'gallery',
						command: 'chart.quick-layout',
						items: chartQuickLayoutGalleryItems,
						sections: chartQuickLayoutGallerySections,
					},
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
				controls: [
					{ kind: 'button', command: 'chart.format-series', size: 'large' },
					{ kind: 'button', command: 'chart.format-area', size: 'large' },
				],
			},
			{
				id: 'chart-delete',
				label: 'Delete',
				controls: [{ kind: 'button', command: 'chart.delete', size: 'large' }],
			},
		],
	};
}
