import type { ToolbarActionId } from 'ooxml-ui/pptx';
import { isPanelVisible, observeElementHeight } from 'ooxml-ui/pptx';
import type { PptxSlide, PptxTextStyleLevels, TextSegment } from 'pptx-viewer-core';
/**
 * ViewerBottomPanels: SlideNotesPanel + StatusBar wrapper.
 *
 * Renders the bottom section of the viewer (notes panel and status bar)
 * when the viewer is not in presentation mode.
 */
import { useEffect, useState } from 'react';

import type { AutosaveStatus } from '../hooks/useAutosave';
import { useToolbarVisibility } from '../hooks/useToolbarVisibility';
import type { ViewerMode } from '../types-core';
import { ResizeHandle } from './ResizeHandle';
import { SlideNotesPanel } from './SlideNotesPanel';
import { StatusBar } from './StatusBar';
import { useViewerCustomizationContext } from './viewer-customization-context';
import { ViewerRegion, useRegionRenderer } from './ViewerRegion';

/* ------------------------------------------------------------------ */
/*  Props                                                             */
/* ------------------------------------------------------------------ */

export interface ViewerBottomPanelsProps {
	activeSlide: PptxSlide | undefined;
	allSlides?: PptxSlide[];
	isSlideNotesCollapsed: boolean;
	canEdit: boolean;
	slideCount: number;
	activeSlideIndex: number;
	isDirty: boolean;
	autosaveStatus?: AutosaveStatus;
	onToggleNotes: () => void;
	onUpdateNotes: (text: string, segments?: TextSegment[]) => void;
	/** Optional collaboration status indicator rendered in the status bar row. */
	collaborationSlot?: React.ReactNode;
	/** Height of the notes panel in pixels (for resizable panels). */
	notesPanelHeight?: number;
	/** Callback to resize the bottom panel. */
	onResizeBottom?: (delta: number) => void;
	/** Current zoom scale. */
	scale?: number;
	/** Zoom in callback. */
	onZoomIn?: () => void;
	/** Zoom out callback. */
	onZoomOut?: () => void;
	/** Zoom to fit callback. */
	onZoomToFit?: () => void;
	/** Current viewer mode. */
	mode?: ViewerMode;
	/** Switch viewer mode. */
	onSetMode?: (mode: ViewerMode) => void;
	/** Toggle slide sorter view. */
	onToggleSlideSorter?: () => void;
	/** When true, the StatusBar row is omitted (mobile uses MobileBottomBar instead). */
	hideStatusBar?: boolean;
	/** Host-supplied list of toolbar buttons/ribbon tabs to hide. */
	hiddenActions?: readonly ToolbarActionId[];
	/**
	 * The deck's notes master `<p:notesStyle>` (`PptxData.notesMaster.
	 * notesStyle`), when loaded. Forwarded to {@link SlideNotesPanel} so the
	 * notes editor and print dialog can fall back to the deck's own
	 * authored notes-text defaults instead of a hardcoded look.
	 */
	notesStyle?: PptxTextStyleLevels;
	/** Live height of the notes strip; see {@link SlideNotesPanel}. Must be stable. */
	onNotesHeightChange?: (height: number) => void;
}

/* ------------------------------------------------------------------ */
/*  Component                                                         */
/* ------------------------------------------------------------------ */

export function ViewerBottomPanels({
	activeSlide,
	allSlides,
	isSlideNotesCollapsed,
	canEdit,
	slideCount,
	activeSlideIndex,
	isDirty,
	autosaveStatus,
	onToggleNotes,
	onUpdateNotes,
	collaborationSlot,
	notesPanelHeight,
	onResizeBottom,
	scale,
	onZoomIn,
	onZoomOut,
	onZoomToFit,
	mode,
	onSetMode,
	onToggleSlideSorter,
	hideStatusBar = false,
	hiddenActions,
	notesStyle,
	onNotesHeightChange,
}: ViewerBottomPanelsProps): React.ReactElement {
	const { isHidden } = useToolbarVisibility(hiddenActions);
	const customization = useViewerCustomizationContext();
	const showNotes = isPanelVisible(customization, 'notes');
	const notesRenderer = useRegionRenderer('notes');
	const statusBarRenderer = useRegionRenderer('statusBar');
	// A replaced notes panel keeps reporting its live height like the built-in
	// one does, so the toast stack still clears it.
	const [notesSlot, setNotesSlot] = useState<HTMLDivElement | null>(null);
	const reportsNotesHeight = Boolean(showNotes && notesRenderer);
	useEffect(() => {
		if (!notesSlot || !reportsNotesHeight || !onNotesHeightChange) {
			return undefined;
		}
		const stop = observeElementHeight(notesSlot, onNotesHeightChange);
		return () => {
			stop();
			onNotesHeightChange(0);
		};
	}, [notesSlot, reportsNotesHeight, onNotesHeightChange]);
	return (
		<>
			{showNotes && onResizeBottom && !isSlideNotesCollapsed && (
				<ResizeHandle direction='vertical' onResize={onResizeBottom} />
			)}
			{showNotes && notesRenderer && (
				// The slot (height, resize handle, collapse) stays the viewer's; a
				// collapsed strip is hidden, not unmounted, so the host keeps its state.
				<div
					ref={setNotesSlot}
					data-pptx-chrome='notes'
					className='flex-shrink-0 overflow-hidden border-t border-border/60 bg-background'
					style={{
						height: notesPanelHeight ?? 120,
						display: isSlideNotesCollapsed ? 'none' : undefined,
					}}
				>
					<ViewerRegion region='notes' renderer={notesRenderer} />
				</div>
			)}
			{showNotes && !notesRenderer && (
				<SlideNotesPanel
					activeSlide={activeSlide}
					allSlides={allSlides}
					isExpanded={!isSlideNotesCollapsed}
					canEdit={canEdit}
					onToggle={onToggleNotes}
					onUpdateNotes={onUpdateNotes}
					panelHeight={notesPanelHeight}
					notesStyle={notesStyle}
					onHeightChange={onNotesHeightChange}
				/>
			)}
			{!hideStatusBar && isPanelVisible(customization, 'statusBar') && statusBarRenderer && (
				<div data-pptx-chrome='status-bar' className='flex-shrink-0'>
					<ViewerRegion region='statusBar' renderer={statusBarRenderer} />
				</div>
			)}
			{!hideStatusBar && isPanelVisible(customization, 'statusBar') && !statusBarRenderer && (
				<StatusBar
					slideCount={slideCount}
					activeSlideIndex={activeSlideIndex}
					isDirty={isDirty}
					autosaveStatus={autosaveStatus}
					scale={scale}
					onZoomIn={onZoomIn}
					onZoomOut={onZoomOut}
					onZoomToFit={onZoomToFit}
					isNotesExpanded={!isSlideNotesCollapsed}
					onToggleNotes={onToggleNotes}
					mode={mode}
					onSetMode={onSetMode as ((mode: 'edit' | 'present') => void) | undefined}
					onToggleSlideSorter={onToggleSlideSorter}
					collaborationSlot={collaborationSlot}
					hideZoomControls={isHidden('zoom')}
					hideNotesToggle={isHidden('notes') || !showNotes}
					hideFullscreenToggle={isHidden('fullscreen')}
				/>
			)}
		</>
	);
}
