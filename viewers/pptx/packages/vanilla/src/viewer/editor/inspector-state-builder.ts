import {
	autoFitModeOf,
	canInteractWithElement,
	hasImageOverrides,
	isElementLocked,
	shouldShowAccessibilitySection,
	textAdvancedStateOf,
	textWrapOf,
} from 'ooxml-ui/pptx';
import type { InlineTextSelection } from 'ooxml-ui/pptx';
/* oxlint-disable eslint/one-var -- pervasive pre-existing pattern in this file
   (many independent short-lived `const`s deriving one inspector field each);
   merging them isn't a style choice here. */
import type { PptxElement, PptxThemeColorRef } from 'pptx-viewer-core';
import {
	hasShapeProperties,
	isImageLikeElement,
	pptxActionToElementAction,
} from 'pptx-viewer-core';

import type { ChartPartSelection } from '../render';
import type { InspectorState } from '../ui';
import { canFormatShape, canFormatText } from './editor-format-mutations';

/** Read the shape fill/stroke for the inspector (undefined/0 when not a shape). */
function shapeStyleOf(el: PptxElement | undefined): {
	fillColor: string | undefined;
	fillColorRef: PptxThemeColorRef | undefined;
	strokeColor: string | undefined;
	strokeColorRef: PptxThemeColorRef | undefined;
	strokeWidth: number;
	fillOpacity: number;
	strokeOpacity: number;
} {
	if (el && hasShapeProperties(el)) {
		return {
			fillColor: el.shapeStyle?.fillColor,
			fillColorRef: el.shapeStyle?.fillColorRef,
			strokeColor: el.shapeStyle?.strokeColor,
			strokeColorRef: el.shapeStyle?.strokeColorRef,
			strokeWidth: el.shapeStyle?.strokeWidth ?? 0,
			fillOpacity: el.shapeStyle?.fillOpacity ?? 1,
			strokeOpacity: el.shapeStyle?.strokeOpacity ?? 1,
		};
	}
	return {
		fillColor: undefined,
		fillColorRef: undefined,
		strokeColor: undefined,
		strokeColorRef: undefined,
		strokeWidth: 0,
		fillOpacity: 1,
		strokeOpacity: 1,
	};
}

/**
 * Build the full {@link InspectorState} for the currently selected element (or
 * the "nothing selected" defaults). Pure; the same element in => the same
 * state out, so it's directly unit-testable without a store/DOM.
 */
export function buildInspectorState(
	el: PptxElement | undefined,
	selectedTableCell: { row: number; column: number } | null = null,
	selectedTableCells: Array<{ row: number; column: number }> = [],
	selectedTextRange: InlineTextSelection | null = null,
	mediaDataUrls: Map<string, string> = new Map(),
	chartPartSelection: ChartPartSelection | null = null,
	recentColors: readonly string[] = [],
	themeColorMap: Record<string, string> | undefined = undefined,
): InspectorState {
	const shape = shapeStyleOf(el);
	const textAdvanced = el ? textAdvancedStateOf(el) : undefined;
	const tableCell =
		el?.type === 'table' && selectedTableCell
			? el.tableData?.rows[selectedTableCell.row]?.cells[selectedTableCell.column]
			: undefined;

	return {
		hasSelection: el !== undefined,
		...(el ? { elementId: el.id } : {}),
		isLocked: isElementLocked(el),
		// G7: a:picLocks/@noCrop. G9: arrowheadsChangeable already existed on
		// element-locks.ts but nothing consulted it in the vanilla inspector.
		croppable: canInteractWithElement(el, 'crop'),
		imageResettable: el !== undefined && isImageLikeElement(el) && hasImageOverrides(el),
		arrowheadsChangeable: canInteractWithElement(el, 'changeArrowheads'),
		canShape: canFormatShape(el),
		canText: canFormatText(el),
		isImage: el !== undefined && isImageLikeElement(el),
		isChart: el?.type === 'chart',
		isMedia: el?.type === 'media',
		isTable: el?.type === 'table',
		isSmartArt: el?.type === 'smartArt',
		smartArtData: el?.type === 'smartArt' ? el.smartArtData : undefined,
		isGroup: el?.type === 'group',
		groupChildCount:
			el?.type === 'group' && Array.isArray(el.children) ? el.children.length : undefined,
		isOle: el?.type === 'ole',
		oleObjectType: el?.type === 'ole' ? el.oleObjectType : undefined,
		oleFileName: el?.type === 'ole' ? el.fileName : undefined,
		oleIsLinked: el?.type === 'ole' && el.isLinked === true,
		oleName: el?.type === 'ole' ? el.oleName : undefined,
		oleElement: el?.type === 'ole' ? el : undefined,
		...(el ? { selectedElement: el } : {}),
		x: el?.x ?? 0,
		y: el?.y ?? 0,
		width: el?.width ?? 0,
		height: el?.height ?? 0,
		rotation: el?.rotation ?? 0,
		fillColor: shape.fillColor,
		fillColorRef: shape.fillColorRef,
		strokeColor: shape.strokeColor,
		strokeColorRef: shape.strokeColorRef,
		strokeWidth: shape.strokeWidth,
		shapeStyle: el && hasShapeProperties(el) ? el.shapeStyle : undefined,
		shapeType: el?.type === 'shape' ? el.shapeType : undefined,
		isConnector: el?.type === 'connector',
		fillOpacity: shape.fillOpacity,
		strokeOpacity: shape.strokeOpacity,
		vAlign: textAdvanced?.vAlign ?? 'top',
		textWrap: el ? textWrapOf(el) : 'square',
		autoFitMode: el ? autoFitModeOf(el) : 'none',
		textDirection: textAdvanced?.textDirection ?? 'horizontal',
		textStyle: el && 'textStyle' in el ? el.textStyle : undefined,
		selectedTextRange,
		imageArtisticEffect:
			el && isImageLikeElement(el) ? (el.imageEffects?.artisticEffect ?? 'none') : 'none',
		imageDuotone1:
			el && isImageLikeElement(el) ? (el.imageEffects?.duotone?.color1 ?? '#000000') : '#000000',
		imageDuotone2:
			el && isImageLikeElement(el) ? (el.imageEffects?.duotone?.color2 ?? '#ffffff') : '#ffffff',
		imageColorWash: el && isImageLikeElement(el) ? el.imageEffects?.colorWash : undefined,
		showAccessibilitySection: el !== undefined && shouldShowAccessibilitySection(el),
		chartHighlightCell:
			el?.type === 'chart' && chartPartSelection?.elementId === el.id
				? {
						seriesIndex: chartPartSelection.part.seriesIndex,
						pointIndex: chartPartSelection.part.pointIndex,
					}
				: null,
		actionClick: el?.actionClick ? pptxActionToElementAction(el.actionClick, 'click') : undefined,
		actionHover: el?.actionHover ? pptxActionToElementAction(el.actionHover, 'hover') : undefined,
		chartData: el?.type === 'chart' ? el.chartData : undefined,
		media: el?.type === 'media' ? el : undefined,
		mediaPreviewUrl:
			el?.type === 'media'
				? (el.mediaData ?? (el.mediaPath ? mediaDataUrls.get(el.mediaPath) : undefined))
				: undefined,
		mediaPosterUrl:
			el?.type === 'media'
				? (el.posterFrameData ??
					(el.posterFramePath ? mediaDataUrls.get(el.posterFramePath) : undefined))
				: undefined,
		tableStyleId: el?.type === 'table' ? (el.tableData?.tableStyleId ?? '') : '',
		tableCellBackground:
			el?.type === 'table'
				? (el.tableData?.rows[0]?.cells[0]?.style?.backgroundColor ?? '#ffffff')
				: '#ffffff',
		tableCellBorder:
			el?.type === 'table'
				? (el.tableData?.rows[0]?.cells[0]?.style?.borderColor ?? '#000000')
				: '#000000',
		selectedTableCell: tableCell ? selectedTableCell : null,
		selectedTableCells: tableCell ? selectedTableCells : [],
		tableCellStyle: tableCell?.style,
		tableColumnWidths: el?.type === 'table' ? (el.tableData?.columnWidths ?? []) : [],
		tableRowHeights:
			el?.type === 'table' ? (el.tableData?.rows.map((row) => row.height ?? 32) ?? []) : [],
		tableElement: el?.type === 'table' ? el : undefined,
		recentColors,
		themeColorMap,
	};
}
