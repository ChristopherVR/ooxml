import * as officeTheme from 'ooxml-ui/theme';
import { ViewerController, createWorkerParser, type VsdxSource } from 'ooxml-ui/visio';
import * as ui from 'ooxml-ui/visio';
/** Host-owned dialogs, as in the PowerPoint building blocks; omitted ones report not-enabled. */
export interface BuildingBlockHostCallbacks {
	onOpenFile?: (() => void) | undefined;
	onOpenPageSetup?: (() => void) | undefined;
	onOpenSettings?: (() => void) | undefined;
	onOpenShareDialog?: (() => void) | undefined;
}
export interface BuildingBlockOptions extends BuildingBlockHostCallbacks {
	content?: VsdxSource | undefined;
	/** Host edit permission. Read-only until it is `true`, so nothing is editable before the host says so. */
	canEdit?: boolean | undefined;
	locale?: string | undefined;
	onDocumentChange?:
		| ((document: NonNullable<ViewerController['state']['document']>) => void)
		| undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
	onDocumentError?: ((error: Error) => void) | undefined;
}
type Api = Pick<
	typeof ui,
	| 'mountDiagramCanvas'
	| 'fitDiagramCanvas'
	| 'mountDiagramInspector'
	| 'mountDiagramShapes'
	| 'diagramToolbarProps'
	| 'diagramMode'
>;
export function buildingBlock<K extends keyof Api>(key: K): Api[K] {
	const value: unknown = Reflect.get(ui, key);
	if (typeof value !== 'function') throw new Error('Update ooxml-ui to use Visio building blocks.');
	return value as Api[K];
}
/** The mode a host sees: 'edit' only when permission is granted and the file can be saved back. */
export function effectiveMode(state: ViewerController['state']): 'edit' | 'view' {
	return buildingBlock('diagramMode')(state);
}
/** Framework-neutral ownership only; each adapter supplies its own reactive state. */
export function createBuildingBlockSession(initial: BuildingBlockOptions = {}) {
	const controller = new ViewerController(createWorkerParser());
	// Apply the host permission before anything can render or be interacted with.
	controller.canEdit = initial.canEdit === true;
	let options: BuildingBlockOptions = initial;
	let source: VsdxSource | undefined;
	let host: HTMLElement | null = null;
	let disposeCanvas: (() => void) | undefined;
	let destroyed = false;
	const stopEvents = controller.onEvent((name, detail) => {
		if (name === 'document-change') {
			const change = detail as {
				document: NonNullable<typeof controller.state.document>;
				dirty: boolean;
			};
			options.onDocumentChange?.(change.document);
			options.onDirtyChange?.(change.dirty);
		} else if (name === 'document-error') options.onDocumentError?.(detail as Error);
	});
	const fit = () => {
		if (host) buildingBlock('fitDiagramCanvas')(host, controller);
	};
	return {
		controller,
		mount(node: HTMLElement | null) {
			if (destroyed || node === host) return;
			disposeCanvas?.();
			disposeCanvas = undefined;
			host = node;
			if (node) disposeCanvas = buildingBlock('mountDiagramCanvas')(node, controller);
		},
		/** Flat control props; pass the framework's latest options so callbacks stay current. */
		toolbar(latest: BuildingBlockOptions = options) {
			return buildingBlock('diagramToolbarProps')(controller, fit, latest.locale, {
				onOpenFile: latest.onOpenFile,
				onOpenPageSetup: latest.onOpenPageSetup,
				onOpenSettings: latest.onOpenSettings,
				onOpenShareDialog: latest.onOpenShareDialog,
			});
		},
		/** Acknowledge a host save: clears dirty and emits the dirty change. */
		markClean() {
			controller.markClean();
		},
		update(next: BuildingBlockOptions) {
			if (destroyed) return;
			options = next;
			controller.canEdit = next.canEdit === true;
			if (source === next.content) return;
			source = next.content;
			controller.cancelLoad();
			if (!source) return;
			const input = source;
			void controller
				.loadSource(async () => {
					const bytes = input instanceof Blob ? new Uint8Array(await input.arrayBuffer()) : input;
					return bytes;
				})
				.catch(() => {
					/* Controller publishes current-source errors through state and events. */
				});
		},
		destroy() {
			if (destroyed) return;
			destroyed = true;
			disposeCanvas?.();
			disposeCanvas = undefined;
			host = null;
			stopEvents();
			controller.destroy();
		},
	};
}
export type BuildingBlockSession = ReturnType<typeof createBuildingBlockSession>;
export type {
	DiagramToolbarProps,
	DiagramToolbarLabels,
	DiagramToolbarHostCallbacks,
	DiagramMode,
	DiagramTool,
	ViewerController,
} from 'ooxml-ui/visio';

/** Resolve the new optional-control theme API only when that control is mounted. */
export function prepareDiagramBlock(host: HTMLElement | null, kind: 'toolbar' | 'tabs') {
	if (!host) return;
	const prepare: unknown = Reflect.get(officeTheme, 'prepareOfficeUiBlock');
	if (typeof prepare !== 'function')
		throw new Error('Update ooxml-ui to theme Visio building blocks.');
	(prepare as (host: HTMLElement, kind: 'toolbar' | 'tabs', product: 'visio') => void)(
		host,
		kind,
		'visio',
	);
}
