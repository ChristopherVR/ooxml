import {
	createBuildingBlockSession,
	emptyBuildingBlockState,
	type BuildingBlockOptions,
	type BuildingBlockState,
} from '../../bindings/src/building-blocks';
import {
	formulaBarRenderer,
	sheetTabsRenderer,
	toolbarRenderer,
	type BlockRenderer,
} from './building-blocks-dom';

export type { BuildingBlockOptions, BuildingBlockState } from '../../bindings/src/building-blocks';
export { toolbarModel, selectReference } from '../../bindings/src/building-blocks';

/**
 * The framework-free building blocks: one workbook engine and a set of optional control mounts,
 * all driven by the same state the framework hooks return. Mount the canvas where the grid goes
 * and any of the controls where your UI wants them; each mount returns its disposer.
 */
export function createViewerBuildingBlocks(initial: BuildingBlockOptions = {}) {
	let state: BuildingBlockState = emptyBuildingBlockState();
	const listeners = new Set<(state: BuildingBlockState) => void>();
	const session = createBuildingBlockSession((next) => {
		state = next;
		for (const listener of [...listeners]) listener(next);
	});
	session.update(initial);
	const mountControl = (host: HTMLElement, create: (host: HTMLElement) => BlockRenderer) => {
		const render = create(host);
		const listener = (next: BuildingBlockState) => render(next.toolbarProps);
		listeners.add(listener);
		render(state.toolbarProps);
		return () => {
			listeners.delete(listener);
			host.replaceChildren();
		};
	};
	return {
		/** The imperative handle (load, save, select, markClean, share...). Throws before mount. */
		handle: session.handle,
		/** The current `toolbarProps`, loading, error, mode, dirty, selection, collaboration, editError. */
		getState: (): BuildingBlockState => state,
		/** Calls back on every state change; returns the unsubscribe function. */
		subscribe(listener: (state: BuildingBlockState) => void): () => void {
			listeners.add(listener);
			return () => {
				listeners.delete(listener);
			};
		},
		/** Applies new options (workbook, bytes, canEdit, callbacks...). */
		update: (options: BuildingBlockOptions) => session.update(options),
		/** Mounts the grid with no application chrome; returns the disposer. */
		mountCanvas(host: HTMLElement): () => void {
			session.mount(host);
			return () => session.mount(null);
		},
		mountToolbar: (host: HTMLElement) => mountControl(host, toolbarRenderer),
		mountFormulaBar: (host: HTMLElement) => mountControl(host, formulaBarRenderer),
		mountSheetTabs: (host: HTMLElement) => mountControl(host, sheetTabsRenderer),
		destroy: () => {
			listeners.clear();
			session.destroy();
		},
	};
}
export type ViewerBuildingBlocks = ReturnType<typeof createViewerBuildingBlocks>;
