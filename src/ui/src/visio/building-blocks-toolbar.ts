import { editErrorMessage, isEditCancellation, visioSelectionIsOnPage } from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController, type ViewerState } from './controller';
import { canvasServices } from './canvas-services';
import { deleteDiagramSelection } from './canvas-keys';
import { DIAGRAM_TOOLS, type DiagramTool } from './diagram-tool';
import { normalizeVisioLocale, translateRibbonText } from './ribbon-locale';

/** Every label a custom toolbar may show; each has an entry in every Visio locale. */
export const DIAGRAM_TOOLBAR_LABEL_KEYS = [
	'Undo',
	'Redo',
	'Fit to Window',
	'Page',
	'Zoom In',
	'Zoom Out',
	'Duplicate',
	'Group',
	'Ungroup',
	'Delete',
	'Cut',
	'Copy',
	'Paste',
	'Find',
	'Previous Result',
	'Next Result',
	'Shapes',
	'Pointer Tool',
	'Rectangle',
	'Ellipse',
	'Line',
	'Freeform',
	'Arc',
	'Pencil',
	'Text',
	'Connector',
	'Open',
	'Page Setup',
	'Options',
	'Share',
] as const;
export type DiagramToolbarLabelKey = (typeof DIAGRAM_TOOLBAR_LABEL_KEYS)[number];
export type DiagramToolbarLabels = Readonly<Record<DiagramToolbarLabelKey, string>>;
/** What the host may do with the diagram right now: 'view' also covers legacy, source-less files. */
export type DiagramMode = 'edit' | 'view';

/**
 * Dialogs are host-owned, as in the PowerPoint building blocks: the library never opens one in
 * surface mode. A workflow whose callback is omitted reports not-enabled (`canOpen...` is false
 * and no `onOpen...` is returned).
 */
export interface DiagramToolbarHostCallbacks {
	onOpenFile?: (() => void) | undefined;
	onOpenPageSetup?: (() => void) | undefined;
	onOpenSettings?: (() => void) | undefined;
	onOpenShareDialog?: (() => void) | undefined;
}
export interface DiagramToolSummary {
	readonly id: DiagramTool;
	readonly label: string;
	readonly active: boolean;
}

/** Whether the diagram can be edited at all: host permission and a source-backed drawing. */
export function diagramMode(state: Pick<ViewerState, 'canEdit' | 'edit'>): DiagramMode {
	return canEditDiagram(state) ? 'edit' : 'view';
}

function labelsFor(locale: string): DiagramToolbarLabels {
	const language = normalizeVisioLocale(locale);
	return Object.fromEntries(
		DIAGRAM_TOOLBAR_LABEL_KEYS.map((label) => [label, translateRibbonText(language, label)]),
	) as DiagramToolbarLabels;
}
const TOOL_LABELS: Readonly<Record<DiagramTool, DiagramToolbarLabelKey>> = {
	pointer: 'Pointer Tool',
	rectangle: 'Rectangle',
	ellipse: 'Ellipse',
	line: 'Line',
	freeform: 'Freeform',
	arc: 'Arc',
	pencil: 'Pencil',
	text: 'Text',
	connector: 'Connector',
};
export function diagramToolbarProps(
	controller: ViewerController,
	fit: () => void,
	locale = 'en',
	callbacks: DiagramToolbarHostCallbacks = {},
) {
	const state = controller.state;
	const labels = labelsFor(locale);
	const idle = !state.loading && !state.edit.busy;
	const canEdit = canEditDiagram(state) && idle;
	const page = state.document?.pages[state.pageIndex];
	const selected = state.selectedShapes.length;
	const onPage =
		!!page && state.selectedShapes.every((shape) => visioSelectionIsOnPage(shape, page.id));
	const hasSource = state.edit.sourceAvailable && idle && !!page;
	// Clipboard and delete failures are announced on the canvas, as in the full viewer.
	const call = (run: () => Promise<void>) => () =>
		run().catch((error: unknown) => {
			if (!isEditCancellation(error)) canvasServices(controller)?.announce(editErrorMessage(error));
		});
	const clipboard = (operation: 'copy' | 'cut' | 'paste') =>
		call(() => canvasServices(controller)?.clipboard(operation) ?? Promise.resolve());
	const tools: DiagramToolSummary[] = (['pointer', ...DIAGRAM_TOOLS] as const).map((id) => ({
		id,
		label: labels[TOOL_LABELS[id]],
		active: state.tool === id,
	}));
	return {
		labels,
		mode: diagramMode(state),
		canEdit,
		pageIndex: state.pageIndex,
		pages: state.document?.pages.map((entry) => ({ id: entry.id, name: entry.name })) ?? [],
		zoom: state.zoom,
		selection: state.selectedShapes,
		canUndo: canEdit && state.edit.canUndo,
		canRedo: canEdit && state.edit.canRedo,
		canDuplicate: canEdit && selected > 0,
		canGroup: canEdit && selected > 1,
		canUngroup: canEdit && selected > 0,
		canDelete: canEdit && selected > 0 && onPage,
		canCut: canEdit && hasSource && selected > 0 && onPage,
		canCopy: hasSource && selected > 0 && onPage,
		canPaste: canEdit && hasSource,
		/** Whether the shape-creation tools and the shapes pane may add shapes. */
		canDraw: canEdit && !!page,
		canZoomIn: !!page && state.zoom < 8,
		canZoomOut: !!page && state.zoom > 0.1,
		tool: state.tool,
		tools,
		search: {
			query: state.search.query,
			resultCount: state.search.results.length,
			activeIndex: state.search.activeIndex,
			truncated: state.search.indexTruncated || state.search.resultsTruncated,
		},
		onPageChange: (index: number) => controller.setPage(index),
		onZoomChange: (zoom: number) => controller.setZoom(zoom),
		onZoomIn: () => controller.setZoom(controller.state.zoom * 1.25),
		onZoomOut: () => controller.setZoom(controller.state.zoom / 1.25),
		onFit: fit,
		onUndo: () => controller.undo(),
		onRedo: () => controller.redo(),
		onDuplicate: () => controller.duplicateSelection(),
		onGroup: () => controller.groupSelection('group'),
		onUngroup: () => controller.groupSelection('ungroup'),
		onDelete: call(() =>
			deleteDiagramSelection(controller, (message) =>
				canvasServices(controller)?.announce(message),
			),
		),
		onCut: clipboard('cut'),
		onCopy: clipboard('copy'),
		onPaste: clipboard('paste'),
		onToolChange: (tool: DiagramTool) => controller.setTool(tool),
		onTextChange: (pageId: string, shapeId: string, text: string) =>
			controller.replacePlainText(pageId, shapeId, text),
		onSearchChange: (query: string) => controller.setSearchQuery(query),
		onNextSearchResult: () => controller.nextSearchResult(),
		onPreviousSearchResult: () => controller.previousSearchResult(),
		canOpenFile: !!callbacks.onOpenFile,
		canOpenPageSetup: !!callbacks.onOpenPageSetup,
		canOpenSettings: !!callbacks.onOpenSettings,
		canOpenShareDialog: !!callbacks.onOpenShareDialog,
		...(callbacks.onOpenFile ? { onOpenFile: callbacks.onOpenFile } : {}),
		...(callbacks.onOpenPageSetup ? { onOpenPageSetup: callbacks.onOpenPageSetup } : {}),
		...(callbacks.onOpenSettings ? { onOpenSettings: callbacks.onOpenSettings } : {}),
		...(callbacks.onOpenShareDialog ? { onOpenShareDialog: callbacks.onOpenShareDialog } : {}),
	};
}
export type DiagramToolbarProps = ReturnType<typeof diagramToolbarProps>;
