import {
	openPptxFile,
	readBackstageRecentFile,
	resolveCollaborationConfig,
	resolveCollaborationShellEditability,
	resolveCollaborationShellState,
} from 'ooxml-ui/pptx';
import type { CollabLoadOrigin } from 'ooxml-ui/pptx';
/**
 * useViewerBuildingBlocks: Composes the same state + hooks `PowerPointViewer`
 * wires internally, and maps them into flat prop objects for the standalone
 * `Toolbar` and `SlideCanvas` components.
 *
 * This lets a host assemble its own viewer shell instead of only getting the
 * bundled `<PowerPointViewer>`:
 *
 * ```tsx
 * function MyCustomViewer({ content }: { content: Uint8Array }) {
 *   const { toolbarProps, canvasProps, slidesPaneProps } = useViewerBuildingBlocks({
 *     content,
 *     canEdit: true,
 *   });
 *   return (
 *     <>
 *       <Toolbar {...toolbarProps} />
 *       <div style={{ display: 'flex' }}>
 *         <SlidesPane {...slidesPaneProps} />
 *         <SlideCanvas {...canvasProps} />
 *       </div>
 *     </>
 *   );
 * }
 * ```
 *
 * This is an additive composition, not a refactor: `PowerPointViewer` keeps
 * its own independent wiring in `ViewerToolbarSection` / `ViewerMainContent`,
 * and this hook intentionally duplicates that wiring (via
 * `useViewerBuildingBlocksState`) rather than routing `PowerPointViewer`
 * through it (a much larger, riskier change). Pieces `PowerPointViewer`
 * renders itself, dialogs, presentation overlays, mobile chrome, resizable
 * panels, are out of scope here; hosts that need them should
 * render `PowerPointViewer` instead.
 */
import { useCallback, useEffect, useMemo, useState } from 'react';

import { useTranslation } from '../i18n/use-translation';
import { buildCanvasProps } from './useViewerBuildingBlocks-canvas-props';
import { useSlidesPaneProps } from './useViewerBuildingBlocks-slides-pane-props';
import { buildToolbarProps } from './useViewerBuildingBlocks-toolbar-props';
import type {
	UseViewerBuildingBlocksInput,
	ViewerBuildingBlocksResult,
} from './useViewerBuildingBlocks-types';
import { useViewerBuildingBlocksCollaboration } from './useViewerBuildingBlocksCollaboration';
import { useViewerBuildingBlocksCore } from './useViewerBuildingBlocksCore';
import { useViewerBuildingBlocksState } from './useViewerBuildingBlocksState';

export type {
	UseViewerBuildingBlocksInput,
	ViewerBuildingBlocksResult,
} from './useViewerBuildingBlocks-types';

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useViewerBuildingBlocks(
	input: UseViewerBuildingBlocksInput,
): ViewerBuildingBlocksResult {
	const {
		content: incomingContent,
		canEdit: requestedCanEdit = false,
		filePath,
		fileName,
		initialSlide,
		googleFonts,
		listConversionScope,
		allowExternalContent,
		limits,
		autosaveEnabled = true,
		userName,
		hiddenActions,
		handle,
		onContentChange,
		onDirtyChange,
		onActiveSlideChange,
		onModeChange,
		onZoomChange,
		onSelectionChange,
		onSlideCountChange,
		onOpenSettings,
		onOpenHeaderFooter,
		onOpenShareDialog,
		onOpenLayoutGallery,
	} = input;
	const collaborationConfig = resolveCollaborationConfig(input.collaboration);
	const [collaborationReadOnly, setCollaborationReadOnly] = useState(Boolean(collaborationConfig));
	const { t } = useTranslation();

	// Local content state, synced from the incoming prop but able to diverge
	// when the built-in File ▸ Open picker loads a different deck in place
	// (mirrors PowerPointViewer's own content state).
	const [content, setContent] = useState<ArrayBuffer | Uint8Array | null>(incomingContent ?? null);
	const [loadOrigin, setLoadOrigin] = useState<CollabLoadOrigin>('bootstrap');
	useEffect(() => {
		setContent(incomingContent ?? null);
		setLoadOrigin('bootstrap');
	}, [incomingContent]);

	const onOpenFile = useCallback(() => {
		void (async () => {
			const picked = await openPptxFile();
			if (picked) {
				setLoadOrigin('user');
				setContent(picked.buffer);
			}
		})();
	}, []);
	const onOpenRecentFile = useCallback((key: string) => {
		void (async () => {
			const bytes = await readBackstageRecentFile(key);
			if (bytes) {
				setLoadOrigin('user');
				setContent(bytes);
			}
		})();
	}, []);

	const core = useViewerBuildingBlocksCore({
		content,
		canEdit: requestedCanEdit,
		fitPadding: input.fitPadding,
		maxFitScale: input.maxFitScale,
	});
	const {
		state,
		mode,
		slides,
		loading,
		error,
		activeSlideIndex,
		activeSlide,
		selectedElement,
		zoom,
		history,
		presentation,
		masterPseudoSlide,
		gridSpacingPx,
		slideSectionGroups,
		viewerOptions,
	} = core;
	// Shared slides can arrive before their original PPTX resources finish loading.
	const canEdit = resolveCollaborationShellEditability({
		authorizedCanEdit: requestedCanEdit,
		configured: Boolean(collaborationConfig),
		readOnly: collaborationReadOnly,
		sourcePending: Boolean(content) && loading,
		sourceError: Boolean(error),
	});

	const {
		dialogs,
		loadVersion,
		editorOps,
		exportHandlers,
		printHandlers,
		propertyHandlers,
		handleSetMode,
		handleEnterPresenterView,
		handleEnterRehearsalMode,
		autosaveStatus,
		viewPreferencesSync,
	} = useViewerBuildingBlocksState({
		core,
		content,
		canEdit,
		filePath,
		initialSlide,
		googleFonts,
		listConversionScope,
		allowExternalContent,
		limits,
		autosaveEnabled,
		userName,
		handle,
		setContent,
		onContentChange,
		onDirtyChange,
		onActiveSlideChange,
		onModeChange,
		onZoomChange,
		onSelectionChange,
		onSlideCountChange,
	});

	// ── Map hook outputs into flat component props ───────────────────────
	const { collaboration, overlay } = useViewerBuildingBlocksCollaboration({
		core,
		config: collaborationConfig,
		content,
		loadOrigin,
		loadVersion,
		embedFonts: dialogs.embedFontsEnabled,
		onReadOnlyChange: setCollaborationReadOnly,
	});
	const toolbarProps = buildToolbarProps({
		mode,
		canEdit,
		state,
		selectedElement,
		activeSlide,
		zoom,
		history,
		findReplace: editorOps.findReplace,
		manipulation: editorOps.manipulation,
		insertHandlers: editorOps.insertHandlers,
		exportHandlers,
		printHandlers,
		propertyHandlers,
		dialogs,
		activeSlideIndex,
		slideOps: editorOps.slideOps,
		ops: editorOps.ops,
		onSetMode: handleSetMode,
		onEnterPresenterView: handleEnterPresenterView,
		onEnterRehearsalMode: handleEnterRehearsalMode,
		onOpenSettings,
		onOpenHeaderFooter,
		onOpenShareDialog,
		onOpenFile,
		onOpenRecentFile,
		fileName,
		autosaveStatus,
		autosaveEnabled,
		hiddenActions,
		viewPreferencesSync,
	});

	const canvasProps = buildCanvasProps({
		mode,
		canEdit,
		slides,
		activeSlide,
		masterPseudoSlide,
		templateElements: state.templateElements,
		canvasSize: state.canvasSize,
		activeSlideIndex,
		gridSpacingPx,
		zoom,
		state,
		selectedElement,
		canvasHandlers: editorOps.canvasHandlers,
		insertHandlers: editorOps.insertHandlers,
		tableOps: editorOps.tableOps,
		presentation,
		findResults: editorOps.findReplace.findResults,
		findResultIndex: editorOps.findReplace.findResultIndex,
		viewerOptions,
		buildHyperlinkConfirmMessage: (url) => `${t('pptx.options.trust.confirmHyperlinks')}\n\n${url}`,
	});

	canvasProps.collaborationOverlay = overlay;
	const slidesPaneProps = useSlidesPaneProps({
		state,
		slides,
		canEdit,
		activeSlideIndex,
		slideSectionGroups,
		slideOps: editorOps.slideOps,
		sectionOps: editorOps.sectionOps,
		recordedTimings: presentation.recordedTimings,
		onOpenLayoutGallery,
	});
	// The same shell projection the other bindings hand their custom chrome.
	const shellState = useMemo(
		() =>
			resolveCollaborationShellState({
				authorizedCanEdit: requestedCanEdit,
				configured: Boolean(collaborationConfig),
				readOnly: collaborationReadOnly,
				sourcePending: Boolean(content) && loading,
				sourceError: Boolean(error),
				status: collaboration?.status ?? 'disconnected',
				remoteUsers: collaboration?.remoteUsers ?? [],
			}),
		[
			requestedCanEdit,
			collaborationConfig,
			collaborationReadOnly,
			content,
			loading,
			error,
			collaboration?.status,
			collaboration?.remoteUsers,
		],
	);
	return {
		toolbarProps,
		canvasProps,
		slidesPaneProps,
		mode,
		loading,
		error,
		autosaveStatus,
		collaboration,
		shellState,
	};
}
