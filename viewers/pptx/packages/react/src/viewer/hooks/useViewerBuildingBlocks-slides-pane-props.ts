/**
 * Builds the flat props for the standalone `<SlidesPane>` (the slide
 * thumbnail rail) from `useViewerBuildingBlocks`'s state, wired the same way
 * `ViewerMainContent` wires the `SlidesPaneSidebar` that `PowerPointViewer`
 * renders.
 *
 * The one difference is the slide context menu's "Layout" entry: in
 * `PowerPointViewer` it opens the Layout gallery, which these building blocks
 * do not render. Here it makes the slide active and hands the request to the
 * host's `onOpenLayoutGallery`, like the toolbar's "Settings" and "Header &
 * Footer" buttons.
 */
import {
	buildSlidesPaneFieldContext,
	buildSlidesPaneTableStyleContext,
	openSlidesPaneLayoutGallery,
	slidesPaneRehearsalTimings,
} from 'ooxml-ui/pptx';
import type { SlidesPaneLayoutGalleryRequest } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { useCallback, useMemo } from 'react';

import type { SlidesPaneSidebarProps } from '../components/slides-pane/types';
import type { SlideSectionGroup } from '../types';
import type { EditorOperationsResult } from './useEditorOperations';
import type { UsePresentationModeResult } from './usePresentationMode';
import type { ViewerState } from './useViewerState';

export type { SlidesPaneLayoutGalleryRequest };

export interface UseSlidesPanePropsInput {
	state: ViewerState;
	slides: PptxSlide[];
	canEdit: boolean;
	activeSlideIndex: number;
	slideSectionGroups: SlideSectionGroup[];
	slideOps: EditorOperationsResult['slideOps'];
	sectionOps: EditorOperationsResult['sectionOps'];
	recordedTimings: UsePresentationModeResult['recordedTimings'];
	onOpenLayoutGallery?: (request: SlidesPaneLayoutGalleryRequest) => void;
}

export function useSlidesPaneProps({
	state,
	slides,
	canEdit,
	activeSlideIndex,
	slideSectionGroups,
	slideOps,
	sectionOps,
	recordedTimings,
	onOpenLayoutGallery,
}: UseSlidesPanePropsInput): SlidesPaneSidebarProps {
	const { setActiveSlideIndex, setIsSlidesPaneOpen } = state;

	// Presentation-wide field/table context so thumbnails substitute fields and
	// resolve table theme colours as the canvas does. Kept stable so the
	// memoised thumbnails do not re-render every frame; each thumbnail re-points
	// the per-slide number and title at its own slide (as `ViewerMainContent`).
	const fieldContext = useMemo(
		() =>
			buildSlidesPaneFieldContext({
				headerFooter: state.headerFooter,
				customProperties: state.customProperties,
				firstSlideNumber: state.firstSlideNumber,
			}),
		[state.headerFooter, state.customProperties, state.firstSlideNumber],
	);
	const tableStyleContext = useMemo(
		() => buildSlidesPaneTableStyleContext(state.theme, state.tableStyleMap),
		[state.theme, state.tableStyleMap],
	);
	const rehearsalTimings = useMemo(
		() => slidesPaneRehearsalTimings(recordedTimings),
		[recordedTimings],
	);

	const onOpenLayoutForSlide = useCallback(
		(slideIndex: number, x: number, y: number) => {
			openSlidesPaneLayoutGallery({ slideIndex, x, y }, setActiveSlideIndex, onOpenLayoutGallery);
		},
		[setActiveSlideIndex, onOpenLayoutGallery],
	);
	const onCollapse = useCallback(() => setIsSlidesPaneOpen(false), [setIsSlidesPaneOpen]);

	return {
		slides,
		templateElementsBySlideId: state.templateElementsBySlideId,
		activeSlideIndex,
		canvasSize: state.canvasSize,
		sectionGroups: slideSectionGroups,
		isOpen: state.isSlidesPaneOpen,
		canEdit,
		onSelectSlide: setActiveSlideIndex,
		onSlideContextMenu: slideOps.handleSlideContextMenu,
		onMoveSlide: slideOps.handleMoveSlide,
		onAddSlide: slideOps.handleAddSlide,
		onAddSlideAfter: slideOps.handleAddSlideAfter,
		onDuplicateSlides: slideOps.handleDuplicateSlides,
		onDeleteSlides: slideOps.handleDeleteSlides,
		onHideSlides: slideOps.handleToggleHideSlides,
		onOpenLayoutForSlide,
		onCollapse,
		onAddSection: sectionOps.addSection,
		onRenameSection: sectionOps.renameSection,
		onDeleteSection: sectionOps.deleteSection,
		onMoveSectionUp: sectionOps.moveSectionUp,
		onMoveSectionDown: sectionOps.moveSectionDown,
		onToggleSectionCollapse: sectionOps.toggleSectionCollapse,
		rehearsalTimings,
		fieldContext,
		tableStyleContext,
	};
}
