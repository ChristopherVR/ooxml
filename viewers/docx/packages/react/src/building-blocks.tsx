import {
	createElement,
	useEffect,
	useImperativeHandle,
	useState,
	useSyncExternalStore,
	type Ref,
} from 'react';
import type { EditorHandle } from '../../bindings/src/index';
import {
	createBuildingBlockSession,
	prepareWordToolbar,
	type BuildingBlockOptions as SessionOptions,
	type CanvasProps,
	type WordNavigationProps,
	type WordStatusProps,
	type WordToolbarProps,
} from '../../bindings/src/building-blocks';
import { navigationControls, statusControls, toolbarControls } from '../../bindings/src/controls';
import { renderControlGroups } from './controls';
export type { CanvasProps } from '../../bindings/src/building-blocks';
export interface BuildingBlockOptions extends SessionOptions {
	handle?: Ref<EditorHandle> | undefined;
}
/** A document surface with no application controls. Keep the returned props together. */
export function DocumentCanvas(props: CanvasProps) {
	return createElement('div', { ref: props.mount, className: props.className });
}
/**
 * Wires the editor to independently rendered pieces: `canvasProps` for `DocumentCanvas`, and flat
 * state with callbacks (`toolbarProps`, `navigationProps`, `statusProps`) for your own controls or
 * the optional `Toolbar`, `PageNavigation` and `StatusBar`. Engine lifetime is the shared session
 * every framework uses; the canvas mounts and releases the engine.
 */
export function useViewerBuildingBlocks(options: BuildingBlockOptions = {}) {
	const [session] = useState(() => createBuildingBlockSession(options));
	const state = useSyncExternalStore(
		session.subscribe,
		() => session.state,
		() => session.state,
	);
	useEffect(() => {
		session.update(options);
	}, [session, options]);
	useImperativeHandle(options.handle, () => session.handle, [session]);
	return {
		canvasProps: session.canvasProps,
		toolbarProps: state.toolbarProps,
		navigationProps: state.navigationProps,
		statusProps: state.statusProps,
		loading: state.loading,
		error: state.error,
		mode: state.mode,
	};
}
export type { WordNavigationProps, WordStatusProps, WordToolbarProps };

/**
 * Optional controls using exactly the callbacks exposed to custom host controls: history, font
 * family and size, formatting, alignment, lists, and the host workflows that are enabled.
 */
export function Toolbar(props: WordToolbarProps) {
	return createElement(
		'div',
		{ role: 'toolbar', ref: prepareWordToolbar },
		...renderControlGroups(toolbarControls(props)),
	);
}

/** Optional page navigation for `navigationProps`. Pages exist in Print Layout only. */
export function PageNavigation(props: WordNavigationProps) {
	return createElement(
		'div',
		{ role: 'group', ref: prepareWordToolbar },
		...renderControlGroups(navigationControls(props)),
	);
}

/** Optional word count and zoom for `statusProps`. */
export function StatusBar(props: WordStatusProps) {
	return createElement(
		'div',
		{ role: 'group', ref: prepareWordToolbar },
		...renderControlGroups(statusControls(props)),
	);
}
