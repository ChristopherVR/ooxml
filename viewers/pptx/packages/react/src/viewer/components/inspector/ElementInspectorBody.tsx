import { sizePositionSections, elementLockTogglePatch, isElementLocked } from 'ooxml-ui/pptx';
import type { FormatSectionView } from 'ooxml-ui/pptx';
import type {
	PptxElement,
	PptxSlide,
	TablePptxElement,
	ChartPptxElement,
	MediaPptxElement,
	ShapeStyle,
	TextStyle,
	ParsedTableStyleMap,
} from 'pptx-viewer-core';
import { isImageLikeElement } from 'pptx-viewer-core';
import React from 'react';
import { LuLock, LuLockOpen } from 'react-icons/lu';

import { useTranslation } from '../../i18n/use-translation';
import type { TableCellEditorState } from '../../types';
import { ActionSettingsPanel } from './ActionSettingsPanel';
import { ChartDataPanel } from './ChartDataPanel';
import {
	ConnectorPanel,
	GroupInfoPanel,
	OlePropertiesPanel,
	TransformPanel,
	LayerOrderButtons,
} from './ElementMiscPanels';
import { FormatSections } from './FormatSections';
import { ImagePropertiesPanel } from './ImagePropertiesPanel';
import { InspectorFields } from './InspectorFields';
import { MediaPropertiesPanel } from './MediaPropertiesPanel';
import { ShapeTextPanels } from './ShapeTextPanels';
import type { ShapeTextPart } from './ShapeTextPanels';
import { SmartArtPropertiesPanel } from './SmartArtPropertiesPanel';
import { TableDataGrid } from './TableDataGrid';
import { TablePropertiesPanel } from './TablePropertiesPanel';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/**
 * Props for the {@link ElementInspectorBody} component.
 */
interface ElementInspectorBodyProps {
	/** The currently selected element to inspect and edit. */
	selectedElement: PptxElement;
	/** Whether editing controls should be enabled. */
	canEdit: boolean;
	/** All slides in the presentation (used by ActionSettingsPanel for hyperlink targets). */
	slides: PptxSlide[];
	/** `data.customShows`, for the Action Settings `customShow` target picker. */
	customShows: Array<{ id: string; name: string }>;
	/** Active table cell editing state, if a table cell is being edited. */
	tableEditorState?: TableCellEditorState | null;
	/** Map of media relationship IDs to data URLs for media preview. */
	mediaDataUrls?: Map<string, string>;
	/**
	 * The deck's parsed `ppt/tableStyles.xml` map, needed by "Edit style...".
	 * See `TablePropertiesPanel`'s docblock for why this is optional.
	 */
	tableStyleMap?: ParsedTableStyleMap;
	onTableStyleMapChange?: (nextMap: ParsedTableStyleMap) => void;
	onDeleteTableStyle?: (styleId: string) => void;
	/** Callback to apply partial updates to the selected element. */
	onUpdateElement: (updates: Partial<PptxElement>) => void;
	/** Callback to apply partial updates to the element's shape style. */
	onUpdateElementStyle: (patch: Partial<ShapeStyle>) => void;
	/** Callback to apply partial updates to the element's text style. */
	onUpdateTextStyle: (patch: Partial<TextStyle>) => void;
	/** Callback to move the element forward or backward in z-order. */
	onMoveLayer: (direction: 'forward' | 'backward') => void;
	/**
	 * Every selected element (the primary one first), for the Format pane's
	 * shared sections and Mixed tags. Edits still go to `selectedElement`.
	 */
	selectedElements?: readonly PptxElement[];
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * Main body of the element inspector panel. Renders type-specific sub-panels
 * based on the selected element type:
 * - Transform controls (position, size, rotation) for all elements
 * - Table properties for table elements
 * - Chart data editing for chart elements
 * - SmartArt properties for SmartArt elements
 * - Image properties for picture/image elements
 * - Media playback controls for media elements
 * - Connector settings for connector elements
 * - Shape text and style panels for shape/text elements
 * - Action settings (hyperlinks, click actions) for actionable elements
 * - Layer ordering controls for all elements
 *
 * @param props - {@link ElementInspectorBodyProps}
 * @returns The composed element inspector body with type-appropriate panels.
 */
export function ElementInspectorBody({
	selectedElement,
	canEdit,
	slides,
	customShows,
	tableEditorState,
	mediaDataUrls,
	tableStyleMap,
	onTableStyleMapChange,
	onDeleteTableStyle,
	onUpdateElement,
	onUpdateElementStyle,
	onUpdateTextStyle,
	onMoveLayer,
	selectedElements,
}: ElementInspectorBodyProps): React.ReactElement {
	const { t } = useTranslation();

	// Shared decides both what reads as "locked" and what the toggle writes, so
	// the button's state can never drift from what the canvas enforces.
	const isLocked = isElementLocked(selectedElement);

	const handleToggleLock = () => {
		if (!canEdit) {
			return;
		}
		onUpdateElement({ locks: elementLockTogglePatch(!isLocked) } as Partial<PptxElement>);
	};

	const shapeText = (part: ShapeTextPart): React.ReactElement => (
		<ShapeTextPanels
			selectedElement={selectedElement}
			canEdit={canEdit}
			onUpdateElement={onUpdateElement}
			onUpdateElementStyle={onUpdateElementStyle}
			onUpdateTextStyle={onUpdateTextStyle}
			part={part}
		/>
	);

	// Each section holds the existing panels `FORMAT_PANEL_SECTIONS` assigns to it.
	const renderSection = (view: FormatSectionView): React.ReactNode => {
		switch (view.id) {
			case 'font':
			case 'paragraph':
			case 'fill':
			case 'effects':
				return shapeText(view.id);
			case 'outline':
				return (
					<>
						{shapeText('outline')}
						{view.arrowheads && (
							<ConnectorPanel
								selectedElement={selectedElement}
								canEdit={canEdit}
								onUpdateElementStyle={onUpdateElementStyle}
							/>
						)}
					</>
				);
			case 'altText':
				return shapeText('altText');
			case 'picture':
				return (
					isImageLikeElement(selectedElement) && (
						<ImagePropertiesPanel
							selectedElement={selectedElement}
							canEdit={canEdit}
							onUpdateElement={onUpdateElement}
						/>
					)
				);
			case 'table':
				return (
					<>
						<TableDataGrid
							tableElement={selectedElement as TablePptxElement}
							canEdit={canEdit}
							onUpdateElement={onUpdateElement}
						/>
						<TablePropertiesPanel
							tableElement={selectedElement as TablePptxElement}
							canEdit={canEdit}
							onUpdateElement={onUpdateElement}
							tableEditorState={tableEditorState}
							tableStyleMap={tableStyleMap}
							onTableStyleMapChange={onTableStyleMapChange}
							onDeleteTableStyle={onDeleteTableStyle}
						/>
					</>
				);
			case 'chart':
				return (
					<ChartDataPanel
						selectedElement={selectedElement as ChartPptxElement}
						canEdit={canEdit}
						onUpdateElement={onUpdateElement}
					/>
				);
			case 'smartArt':
				return (
					selectedElement.type === 'smartArt' &&
					selectedElement.smartArtData && (
						<SmartArtPropertiesPanel
							smartArtData={selectedElement.smartArtData}
							canEdit={canEdit}
							onUpdateElement={onUpdateElement}
							box={{ width: selectedElement.width, height: selectedElement.height }}
						/>
					)
				);
			case 'media':
				return (
					<MediaPropertiesPanel
						element={selectedElement as MediaPptxElement}
						mediaDataUrls={mediaDataUrls ?? new Map()}
						canEdit={canEdit}
						onUpdateElement={onUpdateElement}
					/>
				);
			case 'object':
				return (
					<OlePropertiesPanel
						selectedElement={selectedElement}
						canEdit={canEdit}
						onUpdateElement={onUpdateElement}
					/>
				);
			case 'arrange':
				return (
					<>
						<div className='flex items-center justify-between'>
							<span className='text-muted-foreground'>{t('pptx.inspector.element')}</span>
							<button
								type='button'
								onClick={handleToggleLock}
								disabled={!canEdit}
								title={isLocked ? t('pptx.inspector.unlock') : t('pptx.inspector.lock')}
								className='p-1 rounded hover:bg-accent disabled:opacity-40 disabled:cursor-not-allowed transition-colors max-md:min-w-[44px]! max-md:min-h-[44px]! flex items-center justify-center'
								aria-pressed={isLocked}
							>
								{isLocked ? (
									<LuLock className='w-3.5 h-3.5 text-amber-400' />
								) : (
									<LuLockOpen className='w-3.5 h-3.5 text-muted-foreground' />
								)}
							</button>
						</div>
						{shapeText('shapeType')}
						<GroupInfoPanel selectedElement={selectedElement} />
						<LayerOrderButtons canEdit={canEdit} onMoveLayer={onMoveLayer} />
					</>
				);
			case 'sizePosition':
				return (
					<>
						{/* X / Y / W / H / rotation: the shared Size & Position descriptors. */}
						<div
							className='grid grid-cols-2 gap-1.5 text-[11px]'
							data-pptx-inspector-section='transform'
						>
							<InspectorFields
								sections={sizePositionSections(selectedElement)}
								target={selectedElement}
								onPatch={(patch) => onUpdateElement(patch as Partial<PptxElement>)}
								disabled={!canEdit}
							/>
						</div>
						<TransformPanel
							selectedElement={selectedElement}
							canEdit={canEdit}
							onUpdateElement={onUpdateElement}
						/>
					</>
				);
			case 'action':
				return (
					<ActionSettingsPanel
						selectedElement={selectedElement}
						slides={slides}
						canEdit={canEdit}
						customShows={customShows}
						onUpdateElement={onUpdateElement}
					/>
				);
			default:
				return null;
		}
	};

	return (
		<FormatSections
			elements={selectedElements ?? [selectedElement]}
			renderSection={renderSection}
		/>
	);
}
