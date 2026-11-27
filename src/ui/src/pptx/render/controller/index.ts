/**
 * Headless controllers: framework-free state owners that every binding adapts
 * with one generic store adapter. See `./controller` for the convention.
 *
 * @module render/controller
 */
export { createControllerStore } from './controller';
export type { ControllerStore, PptxController } from './controller';
export {
	createRibbonDisplayController,
	RIBBON_AUTO_COLLAPSE_MEDIA,
} from './ribbon-display-controller';
export type {
	RibbonDisplayController,
	RibbonDisplayControllerOptions,
	RibbonDisplayKey,
	RibbonDisplayState,
} from './ribbon-display-controller';
export { createEditorHistoryController } from './history-controller';
export type {
	EditorHistoryController,
	EditorHistoryControllerOptions,
	EditorHistoryState,
} from './history-controller';
export {
	bindCollaborationHistory,
	collaborationHistoryReady,
	createCollaborationHistory,
} from './collaboration-history';
export type {
	CollaborationHistory,
	CollaborationHistoryReadiness,
	CollaborationHistoryState,
	CollaborationHistoryTarget,
} from './collaboration-history';
export { createAutosaveController } from './autosave-controller';
export type {
	AutosaveController,
	AutosaveControllerDeps,
	AutosaveControllerState,
	AutosaveEngineActivation,
	AutosaveTimers,
	PptxAutosaveStatus,
	PptxAutosaveStatusKind,
} from './autosave-controller';
export { createAutosaveRecoveryController } from './autosave-recovery-controller';
export type {
	AutosaveRecoveryCheckInput,
	AutosaveRecoveryController,
	AutosaveRecoveryControllerDeps,
	AutosaveRecoveryState,
} from './autosave-recovery-controller';
export { createCollaborationSessionController } from './collaboration-session-controller';
export type {
	CollaborationSessionController,
	CollaborationSessionControllerDeps,
	CollaborationSessionState,
	CollaborationSessionTimers,
} from './collaboration-session-controller';
export { createAiAssistantRegistry } from './ai-assistant';
export type {
	AiAssistantRegistry,
	AiAssistantRegistryOptions,
	PptxAiAssistant,
	ResolveAiAssistantOptions,
	ResolvedAiAssistant,
} from './ai-assistant';
export { createCustomShowsController } from './custom-shows-controller';
export type {
	CustomShowsController,
	CustomShowsDialogView,
	CustomShowsResult,
	CustomShowsSlideRow,
	CustomShowsState,
} from './custom-shows-controller';
export {
	SAVE_PROGRESS_TEXT_KEYS,
	createSaveProgressController,
	withSaveProgress,
	saveProgressStatus,
} from './save-progress-controller';
export type {
	SaveProgressController,
	SaveProgressState,
	SaveProgressStatus,
} from './save-progress-controller';
export { createSaveRevisionController } from './save-revision';
export type {
	PptxSaveAcknowledgement,
	PptxSaveSnapshot,
	SaveRevisionController,
	SaveRevisionControllerDeps,
} from './save-revision';
export { createFormatPaneController, formatSectionOpen } from './format-pane-controller';
export type {
	FormatPaneController,
	FormatPaneControllerDeps,
	FormatPaneRequest,
	FormatPaneReveal,
	FormatPaneState,
} from './format-pane-controller';
export { createPresentationRehearsalController } from './presentation-rehearsal-controller';
export type {
	PresentationRehearsalController,
	PresentationRehearsalControllerOptions,
	PresentationRehearsalState,
	RehearsalTimers,
} from './presentation-rehearsal-controller';
export {
	applyPresentationInk,
	createPresentationInkController,
	presentationInkBySlide,
	presentationInkOnSlide,
	presentationInkSlideCount,
	PRESENTATION_ERASER_RADIUS,
	PRESENTATION_HIGHLIGHTER_COLOR,
	PRESENTATION_HIGHLIGHTER_OPACITY,
	PRESENTATION_HIGHLIGHTER_WIDTH,
	PRESENTATION_PEN_COLOR,
	PRESENTATION_PEN_WIDTH,
} from './presentation-ink-controller';
export type {
	PresentationAnnotationStroke,
	PresentationInkController,
	PresentationInkControllerOptions,
	PresentationInkState,
	PresentationInkTimers,
} from './presentation-ink-controller';
export {
	createCommentsPanelController,
	findCommentInThread,
	INITIAL_COMMENTS_PANEL_STATE,
	replaceSlideComments,
} from './comments-panel-controller';
export type {
	CommentsPanelController,
	CommentsPanelControllerDeps,
	CommentsPanelState,
} from './comments-panel-controller';
export { createFindReplaceController } from './find-replace-controller';
export type {
	FindReplaceController,
	FindReplaceControllerState,
	FindReplaceMode,
} from './find-replace-controller';
export {
	createSelectionPaneController,
	formatSelectionPaneLabel,
	SELECTION_PANE_TYPE_KEYS,
	selectionPaneLabel,
	selectionPaneRows,
} from './selection-pane-controller';
export type {
	SelectionPaneController,
	SelectionPaneControllerDeps,
	SelectionPaneLabel,
	SelectionPaneRow,
	SelectionPaneState,
} from './selection-pane-controller';
export { createSlideRailController, slideRailMenuEntries } from './slide-rail-controller';
export type {
	SlideRailClick,
	SlideRailCommands,
	SlideRailController,
	SlideRailControllerDeps,
	SlideRailMenuState,
	SlideRailState,
} from './slide-rail-controller';
export { createSlideSorterController, slideSorterView } from './slide-sorter-controller';
export type {
	SlideSorterClickModifiers,
	SlideSorterCommand,
	SlideSorterController,
	SlideSorterControllerDeps,
	SlideSorterControllerOptions,
	SlideSorterControllerState,
	SlideSorterKeyEvent,
	SlideSorterMenuState,
	SlideSorterOperation,
	SlideSorterView,
} from './slide-sorter-controller';
export { createPresenterSessionController } from './presenter-session-controller';
export type {
	PresenterSessionChannel,
	PresenterSessionController,
	PresenterSessionControllerDeps,
	PresenterSessionRole,
	PresenterSessionState,
} from './presenter-session-controller';
export {
	createProgressiveDeckController,
	decideProgressiveOperation,
	fillPendingSlides,
	isPendingSlide,
	pendingSlide,
	pendingSlideIndex,
	progressiveSearchScope,
	progressiveShowStep,
} from './progressive-deck-controller';
export type {
	ProgressiveDeckController,
	ProgressiveDeckState,
	ProgressiveDecision,
	ProgressiveOperation,
} from './progressive-deck-controller';
