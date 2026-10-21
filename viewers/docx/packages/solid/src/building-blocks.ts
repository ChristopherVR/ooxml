import { createEffect, createSignal, onCleanup, type Accessor } from 'solid-js';
import {
	createBuildingBlockSession,
	prepareWordToolbar,
	type BuildingBlockOptions,
	type CanvasProps,
	type WordNavigationProps,
	type WordStatusProps,
	type WordToolbarProps,
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
	CanvasProps,
	WordNavigationProps,
	WordStatusProps,
	WordToolbarProps,
} from '../../bindings/src/building-blocks';
export function createViewerBuildingBlocks(options: Accessor<BuildingBlockOptions> = () => ({})) {
	const session = createBuildingBlockSession();
	const [state, setState] = createSignal(session.state);
	const stop = session.subscribe(setState);
	createEffect(() => session.update(options()));
	onCleanup(() => {
		stop();
		session.destroy();
	});
	return {
		canvasProps: session.canvasProps,
		handle: session.handle,
		toolbarProps: () => state().toolbarProps,
		navigationProps: () => state().navigationProps,
		statusProps: () => state().statusProps,
		loading: () => state().loading,
		error: () => state().error,
		mode: () => state().mode,
	};
}
export function DocumentCanvas(props: CanvasProps) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	createEffect(() => {
		host.className = props.className ?? '';
	});
	let mounted: CanvasProps['mount'] | undefined;
	createEffect(() => {
		const next = props.mount;
		if (mounted === next) return;
		mounted?.(null);
		mounted = next;
		next(host);
	});
	onCleanup(() => mounted?.(null));
	return host;
}
function controlsHost(role: string, groups: () => ControlGroup[]): HTMLElement | undefined {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	host.setAttribute('role', role);
	prepareWordToolbar(host);
	const view = createControlsView(host);
	createEffect(() => view.update(groups()));
	onCleanup(() => view.destroy());
	return host;
}
/** Optional controls: history, font, formatting, alignment, lists and enabled host workflows. */
export function Toolbar(props: { toolbarProps: WordToolbarProps }) {
	return controlsHost('toolbar', () => toolbarControls(props.toolbarProps));
}
/** Optional page navigation for `navigationProps`. Pages exist in Print Layout only. */
export function PageNavigation(props: { navigationProps: WordNavigationProps }) {
	return controlsHost('group', () => navigationControls(props.navigationProps));
}
/** Optional word count and zoom for `statusProps`. */
export function StatusBar(props: { statusProps: WordStatusProps }) {
	return controlsHost('group', () => statusControls(props.statusProps));
}
