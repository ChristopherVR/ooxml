/**
 * `ooxml-ui/pptx/ai/host`: the part of the AI assistant core with no SDK runtime imports.
 *
 * A viewer needs these helpers even when the host never opens the chat: the deck bridge
 * (`createPptxAiBridge`, focus targets, the deck-data change check, element updates), the
 * panel controller, the canvas change overlay and its animator, and the stored chat history
 * with its log export. The full `ooxml-ui/pptx/ai` barrel also carries the chat session, the
 * tool registry and the server tool set, which reach the optional `ai` SDK, `zod` and the
 * automation schemas. Nothing reachable from this entry imports any of those at runtime
 * (`host.test.ts` checks it), so a binding's main entry can use it without pulling the
 * assistant runtime into the host's initial bundle.
 *
 * This is about runtime imports only. The type declarations still reference SDK types (the
 * stored chat history and log export name the `ai` package's `UIMessage`), so type checking a
 * consumer resolves `ai` types the same way the full barrel does.
 *
 * Every export here is also exported, unchanged, from `ooxml-ui/pptx/ai`.
 */

export { deckDataFieldChanged } from './deck-data-diff';

export { createPptxAiBridge, pptxAiDeckTitle } from './bridge-factory';
export type { PptxAiBridgePort, PptxAiCanvasSize } from './bridge-factory';

export {
	AI_TOOL_FLASH_MS,
	aiFixDirective,
	createPptxAiPanelController,
	isAiFocusPinned,
	resolveAiEffectiveTargets,
} from './panel-controller';
export type {
	PptxAiPanelController,
	PptxAiPanelControllerDeps,
	PptxAiPanelState,
	PptxAiPrefill,
} from './panel-controller';

export type {
	PptxAiBridge,
	PptxAiDataUpdater,
	PptxAiDeckMeta,
	PptxAiElementUpdate,
	PptxAiFocusedTarget,
	PptxAiNotifyLevel,
	PptxAiSlidesUpdater,
} from './bridge';

export {
	computeFocusTargets,
	focusTargetChips,
	isTwoTableFocus,
	mergeTablesDirective,
} from './focus-targets';
export type { AiCanvasHighlight, FocusChip, FocusSelectionInput } from './focus-targets';

export { createAiChangeAnimator, resolveChangeAnimationConfig } from './change-animator';
export type {
	AiChangeAnimationConfig,
	AiChangeAnimator,
	AiChangeBatch,
	ResolvedAiChangeAnimationConfig,
} from './change-animator';
export { diffChangedElements } from './change-diff';
export type { AiChangeBounds, AiChangeKind, AiElementChange } from './change-diff';
export { AI_CHANGE_ATTR, aiChangeAnimationCss, changeGhostStyle } from './change-animation-css';
export type { GhostStyle } from './change-animation-css';

export { createChatHistoryStore } from './chat-history-store';
export type { PptxAiChatStore, PptxAiChatSummary, PptxAiStoredChat } from './chat-history-store';

export {
	buildChatLogExport,
	buildChatLogMarkdown,
	collectStoredChats,
	exportAiChatLogs,
	toLogChat,
	toolCallLine,
} from './chat-log-export';
export type {
	AiLogChat,
	AiLogExport,
	AiLogFormat,
	AiLogMessage,
	AiLogToolCall,
	BuildChatLogOptions,
	ExportAiChatLogsMeta,
	SaveChatLogFile,
} from './chat-log-export';

export { applyElementUpdate, applyShapeStyleUpdate, applyTextUpdate } from './tools/mutations';

export type { ToolCanvasTarget } from './tool-target';
