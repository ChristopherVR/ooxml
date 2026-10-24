import type { DocumentStorage } from 'ooxml-ui/pptx';
import type { PptxCoreProperties, PptxSlide, PptxHandler } from 'pptx-viewer-core';
/**
 * useIOHandlers: Composes export, print, theme, and property-handler hooks
 * into a single grouped return value for the orchestrator component.
 */
import type React from 'react';

import type { CanvasSize } from '../types';
import type { EditorHistoryResult } from './useEditorHistory';
import type { ElementOperations } from './useElementOperations';
import { useExportHandlers } from './useExportHandlers';
import type { ExportHandlersResult } from './useExportHandlers';
import { usePrintHandlers } from './usePrintHandlers';
import type { PrintHandlersResult } from './usePrintHandlers';
import { usePropertyHandlers } from './usePropertyHandlers';
import type { PropertyHandlersResult } from './usePropertyHandlers';
import type { SerializeSlides } from './useSerialize';
import { useThemeHandlers } from './useThemeHandlers';
import type { ThemeHandlersResult } from './useThemeHandlers';
import type { ViewerState } from './useViewerState';

// ---------------------------------------------------------------------------
// Input
// ---------------------------------------------------------------------------

export interface UseIOHandlersInput {
	state: ViewerState;
	slides: PptxSlide[];
	activeSlideIndex: number;
	canvasSize: CanvasSize;
	filePath: string | undefined;
	/** Host `storage` prop, forwarded to the Save / Save As handlers. */
	storage?: DocumentStorage | undefined;
	/** Reports a failed Save / Save As to the host and the user. */
	onSaveError?: ((error: unknown) => void) | undefined;
	history: EditorHistoryResult;
	ops: ElementOperations;
	zoom: {
		canvasStageRef: React.RefObject<HTMLDivElement | null>;
	};
	handlerRef: React.RefObject<PptxHandler | null>;
	/**
	 * The one user-facing serialiser (`useSerialize`). Save As passes it the
	 * output format, so the downloaded file carries every save option Save does.
	 */
	serializeSlides: SerializeSlides;
	setContent: React.Dispatch<React.SetStateAction<ArrayBuffer | Uint8Array | null>>;
	/** The host's `onPropertiesChange` (document-properties edits). */
	onPropertiesChange?: ((updates: Partial<PptxCoreProperties>) => void) | undefined;
	/** The host's `onNotesUpdate` (speaker-notes edits). */
	onNotesUpdate?: ((notes: string) => void) | undefined;
	/** File > Options > Advanced > "Image Size and Quality" raster-scale multiplier. */
	imageExportScale?: number;
	/**
	 * The raw `resolveImageResolutionScale(viewerOptions)` multiplier, fed into
	 * GIF/video export's shared `resolveExportCaptureDecision`. See the
	 * matching field on `UseExportHandlersInput`.
	 */
	imageResolutionScale?: number;
}

// ---------------------------------------------------------------------------
// Output
// ---------------------------------------------------------------------------

export interface IOHandlersResult {
	exportHandlers: ExportHandlersResult;
	printHandlers: PrintHandlersResult;
	themeHandlers: ThemeHandlersResult;
	propertyHandlers: PropertyHandlersResult;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useIOHandlers(input: UseIOHandlersInput): IOHandlersResult {
	const {
		state,
		slides,
		activeSlideIndex,
		canvasSize,
		filePath,
		storage,
		onSaveError,
		history,
		ops,
		zoom,
		handlerRef,
		serializeSlides,
		setContent,
		onPropertiesChange,
		onNotesUpdate,
		imageExportScale,
		imageResolutionScale,
	} = input;

	const exportHandlers = useExportHandlers({
		slides,
		templateElementsBySlideId: state.templateElementsBySlideId,
		activeSlide: slides[activeSlideIndex],
		activeSlideIndex,
		filePath,
		storage,
		onSaveError,
		canvasStageRef: zoom.canvasStageRef,
		setActiveSlideIndex: state.setActiveSlideIndex,
		serializeSlides,
		headerFooter: state.headerFooter,
		presentationProperties: state.presentationProperties,
		customShows: state.customShows,
		sections: state.sections,
		coreProperties: state.coreProperties,
		appProperties: state.appProperties,
		customProperties: state.customProperties,
		tagCollections: state.tagCollections,
		notesMaster: state.notesMaster,
		handoutMaster: state.handoutMaster,
		theme: state.theme,
		canvasSize,
		slideSizeEmu: state.slideSizeEmu,
		imageExportScale,
		imageResolutionScale,
	});

	const printHandlers = usePrintHandlers({
		slides,
		activeSlideIndex,
		canvasStageRef: zoom.canvasStageRef,
		setActiveSlideIndex: state.setActiveSlideIndex,
	});

	const themeHandlers = useThemeHandlers({
		handlerRef,
		setTheme: state.setTheme as unknown as React.Dispatch<
			React.SetStateAction<Record<string, unknown> | null>
		>,
		setSlideMasters: state.setSlideMasters as unknown as React.Dispatch<
			React.SetStateAction<Array<Record<string, unknown>>>
		>,
		slideMasters: state.slideMasters as unknown as Array<Record<string, unknown>>,
		themeOptions: state.themeOptions,
		history,
		slides,
		setSlides: state.setSlides,
		templateElementsBySlideId: state.templateElementsBySlideId,
		setTemplateElementsBySlideId: state.setTemplateElementsBySlideId,
		theme: state.theme,
		bumpHistory: () => state.setPointerCommitNonce((n) => n + 1),
	});

	const propertyHandlers = usePropertyHandlers({
		slides,
		activeSlideIndex,
		canvasSize,
		setContent,
		setPresentationProperties: state.setPresentationProperties,
		setCoreProperties: state.setCoreProperties as unknown as React.Dispatch<
			React.SetStateAction<import('pptx-viewer-core').PptxCoreProperties | null>
		>,
		setAppProperties: state.setAppProperties as unknown as React.Dispatch<
			React.SetStateAction<import('pptx-viewer-core').PptxAppProperties | null>
		>,
		setCustomProperties: state.setCustomProperties,
		setSlides: state.setSlides,
		setIsDirty: state.setIsDirty,
		ops,
		history,
		onPropertiesChange,
		onNotesUpdate,
	});

	return { exportHandlers, printHandlers, themeHandlers, propertyHandlers };
}
