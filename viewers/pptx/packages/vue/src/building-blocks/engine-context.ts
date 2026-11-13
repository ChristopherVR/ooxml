import type { PptxLayoutOption } from 'ooxml-core/pptx';
import type { ReadOnlyRecommendation } from 'ooxml-ui/pptx';
import type {
	ParsedSignature,
	PptxAppProperties,
	ParsedTableStyleMap,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxCustomShow,
	PptxElement,
	PptxHandler,
	PptxHeaderFooter,
	PptxPresentationProperties,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';
import type { ComputedRef, Ref, ShallowRef } from 'vue';

import type { UseCollaborationResult } from '../viewer/composables/collaboration-types';
import type { UseCustomShowsWiringResult } from '../viewer/composables/useCustomShowsWiring';
import type { UseDocumentPropertiesDialogResult } from '../viewer/composables/useDocumentPropertiesDialog';
import type { EditorOperations } from '../viewer/composables/useEditorOperations';
import type { UseElementInsertionResult } from '../viewer/composables/useElementInsertion';
import type { UseFindReplaceResult } from '../viewer/composables/useFindReplace';
import type { UseFontEmbeddingResult } from '../viewer/composables/useFontEmbedding';
import type { UseHyperlinkDialogResult } from '../viewer/composables/useHyperlinkDialog';
import type { UseInsertElementDialogsResult } from '../viewer/composables/useInsertElementDialogs';
import type { UseInspectorDeckActionsResult } from '../viewer/composables/useInspectorDeckActions';
import type { UsePasswordProtectionResult } from '../viewer/composables/usePasswordProtection';
import type { UseSectionOperationsResult } from '../viewer/composables/useSectionOperations';
import type { UseSignatureWorkflowResult } from '../viewer/composables/useSignatureWorkflow';
import type { UseSlideOperationsResult } from '../viewer/composables/useSlideOperations';
import type { UseSlideShowSettingsResult } from '../viewer/composables/useSlideShowSettings';
import type { CanvasSize } from '../viewer/types';

/**
 * What `PowerPointViewer` hands the building-blocks engine: its own refs and composable results,
 * nothing re-derived. Declared here so the viewer's wiring stays one call.
 */
export interface BuildingBlocksEngineContext {
	t: (key: string) => string;
	slides: Ref<PptxSlide[]>;
	mergedSlides: ComputedRef<PptxSlide[]>;
	activeSlide: ComputedRef<PptxSlide | undefined>;
	activeSlideIndex: Ref<number>;
	slideCount: ComputedRef<number>;
	canvasSize: Ref<CanvasSize>;
	mediaDataUrls: Ref<Map<string, string>>;
	loading: Ref<boolean>;
	error: Ref<string | null>;
	handler: ShallowRef<PptxHandler | null>;
	layoutOptions: Ref<readonly PptxLayoutOption[]>;
	selectedElements: ComputedRef<PptxElement[]>;
	ops: EditorOperations;
	insertion: UseElementInsertionResult;
	insertDialogs: UseInsertElementDialogsResult;
	hyperlink: UseHyperlinkDialogResult;
	find: UseFindReplaceResult;
	findOpen: Ref<boolean>;
	/** The viewer's mode (`edit`, `preview`, `present`, `master`). */
	mode: ComputedRef<string>;
	canEdit: ComputedRef<boolean>;
	/** The viewer's own permission without Protected View and a read-only recommendation. */
	canEditBase: () => boolean;
	canInsert: () => boolean;
	presenting: Ref<boolean>;
	showMasterView: Ref<boolean>;
	history: {
		canUndo: ComputedRef<boolean> | Ref<boolean>;
		canRedo: ComputedRef<boolean> | Ref<boolean>;
		undo: () => void;
		redo: () => void;
	};
	isDirty: Ref<boolean>;
	goPrev: () => void;
	goNext: () => void;
	zoomIn: () => void;
	zoomOut: () => void;
	zoomReset: () => void;
	password: UsePasswordProtectionResult;
	readOnlyRecommendation: ComputedRef<Pick<ReadOnlyRecommendation, 'kind'>>;
	modifyVerifier: Ref<unknown>;
	customProperties: Ref<PptxCustomProperty[]>;
	coreProperties: Ref<PptxCoreProperties | undefined>;
	appProperties: Ref<PptxAppProperties | undefined>;
	documentProperties: UseDocumentPropertiesDialogResult;
	deckActions: UseInspectorDeckActionsResult;
	fontEmbedding: UseFontEmbeddingResult;
	signatureWorkflow: UseSignatureWorkflowResult;
	signatures: Ref<ParsedSignature[]>;
	activeContent: Ref<ArrayBuffer | Uint8Array | null | undefined>;
	slideShow: UseSlideShowSettingsResult;
	presentationProperties: Ref<PptxPresentationProperties>;
	customShows: Ref<PptxCustomShow[]>;
	customShowsWiring: UseCustomShowsWiringResult;
	broadcastOpen: Ref<boolean>;
	/** The viewer's collaboration session (shell state and the shared `CollaborationState`). */
	collab: UseCollaborationResult;
	showShortcuts: Ref<boolean>;
	slideOps: UseSlideOperationsResult;
	sectionOps: UseSectionOperationsResult;
	hasSections: ComputedRef<boolean>;
	toggleSlideHidden: (index: number) => void;
	headerFooter: Ref<PptxHeaderFooter | undefined>;
	firstSlideNumber: Ref<number | undefined>;
	theme: Ref<PptxTheme | undefined>;
	tableStyleMap: Ref<ParsedTableStyleMap | undefined>;
}
