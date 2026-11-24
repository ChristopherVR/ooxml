/**
 * useExportHandlers: Export to PNG, PDF, Video, GIF, "Copy slide as image",
 * and "Save As" format handlers.
 *
 * Types live in ./export-handler-types.ts;
 * Save-as logic lives in ./useExportSaveAs.ts.
 */
import {
	choosePdfExportSelection,
	EXPORT_ASSEMBLING_PERCENT,
	EXPORT_DONE_PERCENT,
	isExportAbortError,
	slideProgressPercent,
	slideStatusLabel,
} from 'ooxml-ui/pptx';
import { useState, useRef, useCallback } from 'react';

import { useTranslation } from '../i18n/use-translation';
import { exportSlideAsPng, copySlideToClipboard } from '../utils/export';
import { exportPresentationPdf } from '../utils/export-pdf-vector';
import type { UseExportHandlersInput, ExportHandlersResult } from './export-handler-types';
import { useExportMediaHandlers } from './useExportMediaHandlers';
import { useExportNotesPdf } from './useExportNotesPdf';
import { useExportSaveAs } from './useExportSaveAs';

export type { UseExportHandlersInput, ExportHandlersResult } from './export-handler-types';
export type { ExportModalControls } from './export-handler-types';
export { downloadBlob } from '../utils/dom-helpers';

export function useExportHandlers(input: UseExportHandlersInput): ExportHandlersResult {
	const {
		slides,
		templateElementsBySlideId,
		activeSlide,
		activeSlideIndex,
		filePath,
		storage,
		onSaveError,
		canvasStageRef,
		setActiveSlideIndex,
		serializeSlides,
		headerFooter,
		presentationProperties,
		customShows,
		sections,
		coreProperties,
		appProperties,
		customProperties,
		tagCollections,
		notesMaster,
		handoutMaster,
		theme,
		canvasSize,
		slideSizeEmu,
		imageExportScale,
		imageResolutionScale,
		pdfExportSettings,
	} = input;
	const { t } = useTranslation();
	const backgroundColor = activeSlide?.backgroundColor;

	const [exportModalOpen, setExportModalOpen] = useState(false);
	const [exportModalTitle, setExportModalTitle] = useState('');
	const [exportProgress, setExportProgress] = useState(0);
	const [exportStatusMessage, setExportStatusMessage] = useState('');
	const exportAbortRef = useRef<AbortController | null>(null);
	const handleExportNotesPdf = useExportNotesPdf({
		slides,
		activeSlideIndex,
		canvasStageRef,
		setActiveSlideIndex,
		imageExportScale,
		exportAbortRef,
		setExportModalOpen,
		setExportModalTitle,
		setExportProgress,
		setExportStatusMessage,
	});

	const { handleExportVideo, handleExportGif } = useExportMediaHandlers({
		slides,
		activeSlideIndex,
		canvasStageRef,
		setActiveSlideIndex,
		imageResolutionScale,
		exportAbortRef,
		setExportModalOpen,
		setExportModalTitle,
		setExportProgress,
		setExportStatusMessage,
	});

	const {
		handleExportJson,
		handleSaveAsFormat,
		handleSave,
		handleSaveAsPptx,
		handleSaveAsPpsx,
		handleSaveAsPptm,
		handleSaveAsPpt,
	} = useExportSaveAs({
		slides,
		templateElementsBySlideId,
		filePath,
		storage,
		onSaveError,
		serializeSlides,
		headerFooter,
		presentationProperties,
		customShows,
		sections,
		coreProperties,
		appProperties,
		customProperties,
		tagCollections,
		notesMaster,
		handoutMaster,
		theme,
		canvasSize,
		slideSizeEmu,
	});

	const handleExportPng = useCallback(async () => {
		const stageEl = canvasStageRef.current;
		if (!stageEl) {
			return;
		}
		try {
			await exportSlideAsPng(stageEl, activeSlideIndex, {
				backgroundColor,
				scale: imageExportScale,
			});
		} catch (err) {
			console.error('[PowerPointViewer] PNG export failed:', err);
		}
	}, [canvasStageRef, activeSlideIndex, backgroundColor, imageExportScale]);

	// Read at invocation, so the callback stays stable and never holds a previous deck.
	const pdfInputRef = useRef({
		settings: pdfExportSettings ?? { mode: 'vector' as const, includeHiddenSlides: false },
		slides,
		canvasSize,
		slideSize: slideSizeEmu,
	});
	pdfInputRef.current = {
		settings: pdfExportSettings ?? { mode: 'vector', includeHiddenSlides: false },
		slides,
		canvasSize,
		slideSize: slideSizeEmu,
	};
	// The host callbacks and translator, read at invocation like the deck above.
	const pdfHostRef = useRef({ input, t });
	pdfHostRef.current = { input, t };

	const handleExportPdf = useCallback(async () => {
		if (!canvasStageRef.current) {
			return;
		}
		const host = pdfHostRef.current;
		// File > Export > PDF asks for the range first (vector pages only); Cancel exports nothing.
		const selection = await choosePdfExportSelection(host.input.pdfExportContainerRef?.current, {
			settings: pdfInputRef.current.settings,
			slides: pdfInputRef.current.slides,
			currentSlide: activeSlideIndex,
			translate: host.t,
		});
		if (!selection) {
			return;
		}
		const abortCtrl = new AbortController();
		exportAbortRef.current = abortCtrl;
		setExportModalTitle('Export as PDF');
		setExportStatusMessage('Capturing slides...');
		setExportProgress(0);
		setExportModalOpen(true);
		try {
			await exportPresentationPdf(
				canvasStageRef,
				setActiveSlideIndex,
				activeSlideIndex,
				'presentation.pdf',
				{
					...pdfInputRef.current,
					selection,
					notify: host.input.notifyPdfExport,
					dismissNotice: host.input.dismissPdfExportNotice,
					onApproximated: host.input.onPdfExportReport,
					scale: imageExportScale,
					onProgress: (current, total) => {
						setExportProgress(slideProgressPercent(current, total));
						setExportStatusMessage(slideStatusLabel('Rendering', current, total));
					},
					signal: abortCtrl.signal,
				},
			);
			setExportProgress(EXPORT_ASSEMBLING_PERCENT);
			setExportStatusMessage('Building PDF...');
			await new Promise<void>((r) => {
				setTimeout(r, 100);
			});
			setExportProgress(EXPORT_DONE_PERCENT);
			setExportStatusMessage('Done!');
		} catch (err) {
			if (!isExportAbortError(err)) {
				console.error('[PowerPointViewer] PDF export failed:', err);
			}
		} finally {
			exportAbortRef.current = null;
			setExportModalOpen(false);
		}
	}, [canvasStageRef, setActiveSlideIndex, activeSlideIndex, imageExportScale]);

	const handleCopySlideAsImage = useCallback(async () => {
		const stageEl = canvasStageRef.current;
		if (!stageEl) {
			return;
		}
		try {
			await copySlideToClipboard(stageEl, {
				backgroundColor,
				scale: imageExportScale,
			});
		} catch (err) {
			console.error('[PowerPointViewer] Copy slide as image failed:', err);
		}
	}, [canvasStageRef, backgroundColor, imageExportScale]);

	const handleCancelExport = useCallback(() => {
		exportAbortRef.current?.abort();
		exportAbortRef.current = null;
		setExportModalOpen(false);
		setExportProgress(0);
	}, []);

	return {
		handleExportPng,
		handleExportPdf,
		handleExportNotesPdf,
		handleCopySlideAsImage,
		handleExportVideo,
		handleExportGif,
		handleExportJson,
		handleSaveAsFormat,
		handleSave,
		handleSaveAsPptx,
		handleSaveAsPpsx,
		handleSaveAsPptm,
		handleSaveAsPpt,
		handleCancelExport,
		exportModalOpen,
		exportModalTitle,
		exportProgress,
		exportStatusMessage,
	};
}
