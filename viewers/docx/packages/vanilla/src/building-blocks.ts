import {
	createBuildingBlockSession,
	prepareWordToolbar,
	type BuildingBlockState,
} from '../../bindings/src/building-blocks';
import {
	navigationControls,
	statusControls,
	toolbarControls,
	type ControlGroup,
} from '../../bindings/src/controls';
import { createControlsView } from '../../bindings/src/controls-dom';

export type {
	BuildingBlockOptions,
	BuildingBlockState,
	CanvasProps,
} from '../../bindings/src/building-blocks';

/**
 * The engine lifetime shared by every framework: `canvasProps.mount(host)` attaches the document
 * surface, `subscribe` delivers `toolbarProps`, `navigationProps`, `statusProps`, `loading`, `error`
 * and `mode`, `update(options)` changes options and `destroy()` releases the engine.
 */
export const createViewerBuildingBlocks = createBuildingBlockSession;
export type ViewerBuildingBlocks = ReturnType<typeof createViewerBuildingBlocks>;

type Blocks = Pick<ViewerBuildingBlocks, 'subscribe'>;

/** Mounts the document surface (no title bar, ribbon, status bar or dialogs) into `host`. */
export function mountDocumentCanvas(
	host: HTMLElement,
	blocks: Pick<ViewerBuildingBlocks, 'canvasProps'>,
): () => void {
	blocks.canvasProps.mount(host);
	return () => blocks.canvasProps.mount(null);
}

function mountControls(
	host: HTMLElement,
	blocks: Blocks,
	role: string,
	groups: (state: BuildingBlockState) => ControlGroup[] | null,
): () => void {
	prepareWordToolbar(host);
	host.setAttribute('role', role);
	const view = createControlsView(host);
	const stop = blocks.subscribe((state) => view.update(groups(state) ?? []));
	return () => {
		stop();
		view.destroy();
	};
}

/** Optional library toolbar for `toolbarProps`: history, font, formatting, alignment, lists. */
export function mountToolbar(host: HTMLElement, blocks: Blocks): () => void {
	return mountControls(host, blocks, 'toolbar', (state) =>
		state.toolbarProps ? toolbarControls(state.toolbarProps) : null,
	);
}

/** Optional page navigation for `navigationProps`: previous, page n of m, next and layout view. */
export function mountPageNavigation(host: HTMLElement, blocks: Blocks): () => void {
	return mountControls(host, blocks, 'group', (state) =>
		state.navigationProps ? navigationControls(state.navigationProps) : null,
	);
}

/** Optional status bar for `statusProps`: word count and zoom. */
export function mountStatusBar(host: HTMLElement, blocks: Blocks): () => void {
	return mountControls(host, blocks, 'group', (state) =>
		state.statusProps ? statusControls(state.statusProps) : null,
	);
}
