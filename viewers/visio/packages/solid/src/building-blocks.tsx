import { createEffect, createMemo, createSignal, onCleanup, For, type Accessor } from 'solid-js';
import {
	buildingBlock,
	prepareDiagramBlock,
	createBuildingBlockSession,
	type BuildingBlockOptions,
	type DiagramToolbarProps,
	type ViewerController,
} from '../../bindings/src/building-blocks';
export function createViewerBuildingBlocks(options: Accessor<BuildingBlockOptions> = () => ({})) {
	const session = createBuildingBlockSession();
	const [state, setState] = createSignal(session.controller.state);
	const stop = session.controller.subscribe((value) => setState(() => value));
	createEffect(() => session.update(options()));
	onCleanup(() => {
		stop();
		session.destroy();
	});
	const toolbarProps = createMemo(() => {
		state();
		return session.toolbar(options().locale);
	});
	return {
		controller: session.controller,
		state,
		canvasProps: { mount: session.mount },
		toolbarProps,
		inspectorProps: { controller: session.controller },
		loading: () => state().loading,
		error: () => state().error ?? state().edit.error,
		mode: () => (state().canEdit ? ('edit' as const) : ('view' as const)),
	};
}
export function DiagramCanvas(props: {
	mount(node: HTMLElement | null): void;
	class?: string | undefined;
}) {
	let host!: HTMLDivElement;
	createEffect(() => {
		const mount = props.mount;
		mount(host);
		onCleanup(() => mount(null));
	});
	return <div ref={host} class={props.class} />;
}
export function Toolbar(props: DiagramToolbarProps) {
	return (
		<div role="toolbar" ref={(host) => prepareDiagramBlock(host, 'toolbar')}>
			<button type="button" disabled={!props.canUndo} onClick={() => void props.onUndo()}>
				{props.labels.Undo}
			</button>
			<button type="button" disabled={!props.canRedo} onClick={() => void props.onRedo()}>
				{props.labels.Redo}
			</button>
			<button type="button" onClick={props.onFit}>
				{props.labels['Fit to Window']}
			</button>
		</div>
	);
}
export function PageNavigation(props: DiagramToolbarProps) {
	return (
		<select
			ref={(host) => prepareDiagramBlock(host, 'tabs')}
			aria-label={props.labels.Page}
			value={props.pageIndex}
			onChange={(event) => props.onPageChange(Number(event.currentTarget.value))}
		>
			<For each={props.pages}>{(page, index) => <option value={index()}>{page.name}</option>}</For>
		</select>
	);
}
export function Inspector(props: { controller: ViewerController }) {
	let host!: HTMLDivElement;
	createEffect(() => {
		const stop = buildingBlock('mountDiagramInspector')(host, props.controller);
		onCleanup(stop);
	});
	return <div ref={host} />;
}
export type { BuildingBlockOptions, DiagramToolbarProps } from '../../bindings/src/building-blocks';
