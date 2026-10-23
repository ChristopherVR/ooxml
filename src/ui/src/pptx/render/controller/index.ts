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
