/**
 * Headless controllers: framework-free state owners that every binding adapts
 * with one generic store adapter. See `./controller` for the convention.
 *
 * @module render/controller
 */
export { createControllerStore } from './controller';
export type { ControllerStore, PptxController } from './controller';
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
