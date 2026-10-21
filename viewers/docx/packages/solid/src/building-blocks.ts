import { createEffect, createSignal, onCleanup, type Accessor } from 'solid-js';
import {
	createBuildingBlockSession,
	prepareWordToolbar,
	type BuildingBlockOptions,
	type CanvasProps,
	type WordToolbarProps,
} from '../../bindings/src/building-blocks';
export type {
	BuildingBlockOptions,
	CanvasProps,
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
export function Toolbar(props: { toolbarProps: WordToolbarProps }) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	host.setAttribute('role', 'toolbar');
	prepareWordToolbar(host);
	for (const [label, field, action] of [
		['Bold', 'bold', 'onBold'],
		['Italic', 'italic', 'onItalic'],
		['Underline', 'underline', 'onUnderline'],
		['Undo', 'canUndo', 'onUndo'],
		['Redo', 'canRedo', 'onRedo'],
	] as const) {
		const button = document.createElement('button');
		button.type = 'button';
		button.onmousedown = (event) => event.preventDefault();
		button.onclick = () => {
			void props.toolbarProps[action]();
		};
		createEffect(() => {
			const toolbar = props.toolbarProps;
			button.textContent = toolbar.labels[label] ?? label;
			button.disabled =
				action === 'onUndo'
					? !toolbar.canUndo
					: action === 'onRedo'
						? !toolbar.canRedo
						: !toolbar.canEdit;
			button.setAttribute(
				'aria-pressed',
				toolbar[field] === null ? 'mixed' : String(toolbar[field]),
			);
		});
		host.append(button);
	}
	return host;
}
