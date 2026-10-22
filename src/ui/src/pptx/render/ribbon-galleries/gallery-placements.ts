/**
 * Where each gallery sits and which contextual tabs a selection shows; bindings only render it.
 *
 * @module render/ribbon-galleries/gallery-placements
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import type { ResolvedCustomization } from '../customization/customization-resolve';
import type { RibbonControlId, RibbonGroupId } from '../customization/ribbon-control-ids';
import { RIBBON_CONTEXTUAL_TABS } from '../toolbar-actions';
import type { RibbonContextualTabId } from '../toolbar-actions';
import type { RibbonGalleryId } from './gallery-types';

/** One gallery inside a group: the gallery and the control id it is tagged with. */
export interface RibbonGalleryPlacement {
	gallery: RibbonGalleryId;
	control: RibbonControlId;
	/**
	 * `inline`: a strip of the first tiles in the ribbon plus a "more" button
	 * (PowerPoint's in-ribbon galleries). `dropdown`: a single trigger button.
	 */
	mode: 'inline' | 'dropdown';
}

export interface RibbonGalleryGroupPlacement {
	group: RibbonGroupId;
	/** i18n key of the group caption. */
	labelKey: string;
	label: string;
	galleries: readonly RibbonGalleryPlacement[];
	/** The `pptx-ui-ribbon-home-<strip>` drawn after the galleries (Picture Format > Size). */
	strip?: 'picture-size' | 'chart-selection';
}

/** Galleries the format inspector shows as panels (`mode="panel"`) instead of the ribbon. */
export const INSPECTOR_GALLERIES: readonly RibbonGalleryId[] = ['textWarp', 'quickStyles'];

/** The groups each contextual tab renders, in order. */
export const CONTEXTUAL_TAB_GROUPS: Record<
	RibbonContextualTabId,
	readonly RibbonGalleryGroupPlacement[]
> = {
	shapeFormat: [
		{
			group: 'shapeFormat.shapeStyles',
			labelKey: 'pptx.ribbon.groupShapeStyles',
			label: 'Shape Styles',
			galleries: [
				{ gallery: 'shapeStyles', control: 'shapeFormat.shapeStyles.gallery', mode: 'inline' },
				{
					gallery: 'shapeEffects',
					control: 'shapeFormat.shapeStyles.shapeEffects',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'shapeFormat.wordArtStyles',
			labelKey: 'pptx.ribbon.groupWordArtStyles',
			label: 'WordArt Styles',
			galleries: [
				{ gallery: 'wordArtStyles', control: 'shapeFormat.wordArtStyles.gallery', mode: 'inline' },
			],
		},
	],
	pictureFormat: [
		{
			group: 'pictureFormat.adjust',
			labelKey: 'pptx.ribbon.groupAdjust',
			label: 'Adjust',
			galleries: [
				{
					gallery: 'pictureCorrections',
					control: 'pictureFormat.adjust.corrections',
					mode: 'dropdown',
				},
				{ gallery: 'pictureColor', control: 'pictureFormat.adjust.color', mode: 'dropdown' },
				{
					gallery: 'pictureArtisticEffects',
					control: 'pictureFormat.adjust.artisticEffects',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'pictureFormat.pictureStyles',
			labelKey: 'pptx.ribbon.groupPictureStyles',
			label: 'Picture Styles',
			galleries: [
				{
					gallery: 'pictureStyles',
					control: 'pictureFormat.pictureStyles.gallery',
					mode: 'inline',
				},
				{
					gallery: 'shapeEffects',
					control: 'pictureFormat.pictureStyles.pictureEffects',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'pictureFormat.size',
			labelKey: 'pptx.ribbon.groupSize',
			label: 'Size',
			galleries: [],
			strip: 'picture-size',
		},
	],
	tableDesign: [
		{
			group: 'tableDesign.tableStyles',
			labelKey: 'pptx.ribbon.groupTableStyles',
			label: 'Table Styles',
			galleries: [
				{ gallery: 'tableStyles', control: 'tableDesign.tableStyles.gallery', mode: 'inline' },
			],
		},
	],
	tableLayout: [
		{
			group: 'tableLayout.table',
			labelKey: 'pptx.ribbon.groupTable',
			label: 'Table',
			galleries: [
				{ gallery: 'tableSelect', control: 'tableLayout.table.select', mode: 'dropdown' },
				{
					gallery: 'tableViewGridlines',
					control: 'tableLayout.table.viewGridlines',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'tableLayout.rowsColumns',
			labelKey: 'pptx.ribbon.groupRowsColumns',
			label: 'Rows & Columns',
			galleries: [
				{
					gallery: 'tableDeleteRows',
					control: 'tableLayout.rowsColumns.deleteRows',
					mode: 'dropdown',
				},
				{
					gallery: 'tableDeleteColumns',
					control: 'tableLayout.rowsColumns.deleteColumns',
					mode: 'dropdown',
				},
				{
					gallery: 'tableInsertAbove',
					control: 'tableLayout.rowsColumns.insertAbove',
					mode: 'dropdown',
				},
				{
					gallery: 'tableInsertBelow',
					control: 'tableLayout.rowsColumns.insertBelow',
					mode: 'dropdown',
				},
				{
					gallery: 'tableInsertLeft',
					control: 'tableLayout.rowsColumns.insertLeft',
					mode: 'dropdown',
				},
				{
					gallery: 'tableInsertRight',
					control: 'tableLayout.rowsColumns.insertRight',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'tableLayout.merge',
			labelKey: 'pptx.ribbon.groupMerge',
			label: 'Merge',
			galleries: [
				{ gallery: 'tableMergeCells', control: 'tableLayout.merge.mergeCells', mode: 'dropdown' },
				{ gallery: 'tableSplitCells', control: 'tableLayout.merge.splitCells', mode: 'dropdown' },
			],
		},
		{
			group: 'tableLayout.cellSize',
			labelKey: 'pptx.ribbon.groupCellSize',
			label: 'Cell Size',
			galleries: [
				{
					gallery: 'tableDistributeRows',
					control: 'tableLayout.cellSize.distributeRows',
					mode: 'dropdown',
				},
				{
					gallery: 'tableDistributeColumns',
					control: 'tableLayout.cellSize.distributeColumns',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'tableLayout.alignment',
			labelKey: 'pptx.ribbon.groupAlignment',
			label: 'Alignment',
			galleries: [
				{ gallery: 'tableAlignLeft', control: 'tableLayout.alignment.alignLeft', mode: 'dropdown' },
				{
					gallery: 'tableAlignCenter',
					control: 'tableLayout.alignment.alignCenter',
					mode: 'dropdown',
				},
				{
					gallery: 'tableAlignRight',
					control: 'tableLayout.alignment.alignRight',
					mode: 'dropdown',
				},
				{ gallery: 'tableAlignTop', control: 'tableLayout.alignment.alignTop', mode: 'dropdown' },
				{
					gallery: 'tableAlignMiddle',
					control: 'tableLayout.alignment.alignMiddle',
					mode: 'dropdown',
				},
				{
					gallery: 'tableAlignBottom',
					control: 'tableLayout.alignment.alignBottom',
					mode: 'dropdown',
				},
				{
					gallery: 'tableTextDirection',
					control: 'tableLayout.alignment.textDirection',
					mode: 'dropdown',
				},
				{
					gallery: 'tableCellMargins',
					control: 'tableLayout.alignment.cellMargins',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'tableLayout.tableSize',
			labelKey: 'pptx.ribbon.groupTableSize',
			label: 'Table Size',
			galleries: [
				{ gallery: 'tableSize', control: 'tableLayout.tableSize.size', mode: 'dropdown' },
			],
		},
		{
			group: 'tableLayout.arrange',
			labelKey: 'pptx.ribbon.groupArrange',
			label: 'Arrange',
			galleries: [
				{ gallery: 'tableArrange', control: 'tableLayout.arrange.arrange', mode: 'dropdown' },
			],
		},
	],
	chartDesign: [
		{
			group: 'chartDesign.chartLayouts',
			labelKey: 'pptx.ribbon.groupChartLayouts',
			label: 'Chart Layouts',
			galleries: [
				{
					gallery: 'chartQuickLayout',
					control: 'chartDesign.chartLayouts.quickLayout',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'chartDesign.chartStyles',
			labelKey: 'pptx.ribbon.groupChartStyles',
			label: 'Chart Styles',
			galleries: [
				{
					gallery: 'chartColors',
					control: 'chartDesign.chartStyles.changeColors',
					mode: 'dropdown',
				},
				{ gallery: 'chartStyles', control: 'chartDesign.chartStyles.gallery', mode: 'inline' },
			],
		},
	],
	chartFormat: [
		{
			group: 'chartFormat.currentSelection',
			labelKey: 'pptx.ribbon.groupCurrentSelection',
			label: 'Current Selection',
			galleries: [],
			strip: 'chart-selection',
		},
	],
	smartArtDesign: [
		{
			group: 'smartArtDesign.createGraphic',
			labelKey: 'pptx.ribbon.groupCreateGraphic',
			label: 'Create Graphic',
			galleries: [
				{
					gallery: 'smartArtAddShape',
					control: 'smartArtDesign.createGraphic.addShape',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtAddBullet',
					control: 'smartArtDesign.createGraphic.addBullet',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtTextPane',
					control: 'smartArtDesign.createGraphic.textPane',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtPromote',
					control: 'smartArtDesign.createGraphic.promote',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtDemote',
					control: 'smartArtDesign.createGraphic.demote',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtRightToLeft',
					control: 'smartArtDesign.createGraphic.rightToLeft',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtMoveUp',
					control: 'smartArtDesign.createGraphic.moveUp',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtMoveDown',
					control: 'smartArtDesign.createGraphic.moveDown',
					mode: 'dropdown',
				},
			],
		},
		{
			group: 'smartArtDesign.layouts',
			labelKey: 'pptx.ribbon.groupLayouts',
			label: 'Layouts',
			galleries: [
				{ gallery: 'smartArtLayouts', control: 'smartArtDesign.layouts.gallery', mode: 'inline' },
			],
		},
		{
			group: 'smartArtDesign.smartArtStyles',
			labelKey: 'pptx.ribbon.groupSmartArtStyles',
			label: 'SmartArt Styles',
			galleries: [
				{
					gallery: 'smartArtColors',
					control: 'smartArtDesign.smartArtStyles.changeColors',
					mode: 'dropdown',
				},
				{
					gallery: 'smartArtStyles',
					control: 'smartArtDesign.smartArtStyles.gallery',
					mode: 'inline',
				},
			],
		},
		{
			group: 'smartArtDesign.reset',
			labelKey: 'pptx.ribbon.groupReset',
			label: 'Reset',
			galleries: [
				{
					gallery: 'smartArtResetGraphic',
					control: 'smartArtDesign.reset.resetGraphic',
					mode: 'dropdown',
				},
				{ gallery: 'smartArtConvert', control: 'smartArtDesign.reset.convert', mode: 'dropdown' },
			],
		},
	],
};

/**
 * Galleries that live on the fixed tabs, as dropdown triggers inside groups
 * those tabs already render: Home > Drawing (Quick Styles, Shape Effects,
 * replacing the old disabled Shape Effects placeholder), Home > Paragraph
 * (the Bullets / Numbering libraries, next to their toggles) and Design >
 * Variants (a new group after Themes).
 */
export const FIXED_TAB_GALLERIES: readonly RibbonGalleryPlacement[] = [
	{ gallery: 'shapeStyles', control: 'home.drawing.quickStyles', mode: 'dropdown' },
	{ gallery: 'shapeEffects', control: 'home.drawing.shapeEffects', mode: 'dropdown' },
	{ gallery: 'bullets', control: 'home.paragraph.bullets', mode: 'dropdown' },
	{ gallery: 'numbering', control: 'home.paragraph.numbering', mode: 'dropdown' },
	{ gallery: 'themes', control: 'design.themes.gallery', mode: 'inline' },
	{ gallery: 'themeVariants', control: 'design.variants.gallery', mode: 'inline' },
	{ gallery: 'themeColors', control: 'design.variants.colors', mode: 'dropdown' },
	{ gallery: 'themeFonts', control: 'design.variants.fonts', mode: 'dropdown' },
	{ gallery: 'themeEffects', control: 'design.variants.effects', mode: 'dropdown' },
	{
		gallery: 'themeBackgrounds',
		control: 'design.variants.backgroundStyles',
		mode: 'dropdown',
	},
];

/** The contextual tabs a selection of `element` brings up, in PowerPoint's order. */
export function contextualTabsForElement(element: PptxElement | null): RibbonContextualTabId[] {
	if (!element) {
		return [];
	}
	switch (element.type) {
		case 'image':
		case 'picture':
			return ['pictureFormat'];
		case 'table':
			return ['tableDesign', 'tableLayout'];
		case 'chart':
			return ['chartDesign', 'chartFormat'];
		case 'smartArt':
			return ['smartArtDesign'];
		default:
			return hasShapeProperties(element) ? ['shapeFormat'] : [];
	}
}

/** {@link contextualTabsForElement} minus the tabs the host hid. */
export function visibleContextualTabs(
	element: PptxElement | null,
	resolved?: ResolvedCustomization,
): RibbonContextualTabId[] {
	return contextualTabsForElement(element).filter(
		(tab) => !resolved?.hiddenContextualTabs.has(tab),
	);
}

/** i18n key of a contextual tab's caption. */
export function contextualTabLabelKey(tab: RibbonContextualTabId): string {
	return RIBBON_CONTEXTUAL_TABS.find((t) => t.id === tab)?.labelKey ?? tab;
}

/**
 * The tab to show: `active` unless it is a contextual tab the selection no
 * longer brings up, in which case the ribbon falls back to Home (as
 * PowerPoint does when you deselect a picture while on Picture Format).
 */
export function resolveActiveRibbonTab<T extends string>(
	active: T | RibbonContextualTabId,
	visibleContextual: readonly RibbonContextualTabId[],
	fallback: T,
): T | RibbonContextualTabId {
	const isContextual = RIBBON_CONTEXTUAL_TABS.some((t) => t.id === active);
	if (isContextual && !visibleContextual.includes(active as RibbonContextualTabId)) {
		return fallback;
	}
	return active;
}
