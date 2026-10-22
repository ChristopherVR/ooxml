import {
	createEffect,
	createMemo,
	createSignal,
	onCleanup,
	For,
	Index,
	Match,
	Switch,
	type Accessor,
} from 'solid-js';
import {
	buildingBlock,
	prepareDiagramBlock,
	createBuildingBlockSession,
	effectiveMode,
	type BuildingBlockOptions,
	type DiagramToolbarProps,
	type ViewerController,
} from '../../bindings/src/building-blocks';
import {
	toolbarGroups,
	type ToolbarButton,
	type ToolbarControl,
	type ToolbarSearch,
	type ToolbarText,
	type ToolbarGroupId,
} from '../../bindings/src/toolbar-controls';
export function createViewerBuildingBlocks(options: Accessor<BuildingBlockOptions> = () => ({})) {
	// Permission is applied at creation, before the canvas can mount or be used.
	const session = createBuildingBlockSession(options());
	const [state, setState] = createSignal(session.controller.state);
	const stop = session.controller.subscribe((value) => setState(() => value));
	createEffect(() => session.update(options()));
	onCleanup(() => {
		stop();
		session.destroy();
	});
	const toolbarProps = createMemo(() => {
		state();
		return session.toolbar(options());
	});
	return {
		controller: session.controller,
		state,
		canvasProps: { mount: session.mount },
		toolbarProps,
		inspectorProps: { controller: session.controller },
		shapesProps: {
			controller: session.controller,
			get locale() {
				return options().locale;
			},
		},
		/** Acknowledge a host save: clears dirty and fires `onDirtyChange(false)`. */
		markClean: () => session.markClean(),
		loading: () => state().loading,
		error: () => state().error ?? state().edit.error,
		/** Effective permission: host `canEdit` and a drawing that can be saved back. */
		mode: () => effectiveMode(state()),
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
/** Index-keyed, so the DOM (and focus) survives every state update; only attributes change. */
function Control(props: { control: ToolbarControl }) {
	const kind = () => props.control.kind;
	const asButton = () => props.control as ToolbarButton;
	const asSearch = () => props.control as ToolbarSearch;
	return (
		<Switch>
			<Match when={kind() === 'text'}>
				<span data-control={props.control.id}>{(props.control as ToolbarText).text}</span>
			</Match>
			<Match when={kind() === 'button'}>
				<button
					type="button"
					data-control={props.control.id}
					disabled={asButton().disabled}
					aria-pressed={asButton().pressed}
					onClick={() => asButton().run()}
				>
					{asButton().label}
				</button>
			</Match>
			<Match when={kind() === 'search'}>
				<span data-control="search">
					<input
						type="search"
						aria-label={asSearch().label}
						placeholder={asSearch().label}
						value={asSearch().query}
						onInput={(event) => asSearch().change(event.currentTarget.value)}
					/>
					<span aria-live="polite">{asSearch().summary}</span>
					<Control control={asSearch().previous} />
					<Control control={asSearch().next} />
				</span>
			</Match>
		</Switch>
	);
}
/** Optional library toolbar; `include` picks groups. Hosts may build their own from toolbarProps. */
export function Toolbar(props: DiagramToolbarProps & { include?: readonly ToolbarGroupId[] }) {
	const groups = createMemo(() => toolbarGroups(props, props.include));
	return (
		<div role="toolbar" ref={(host) => prepareDiagramBlock(host, 'toolbar')}>
			<Index each={groups()}>
				{(group) => (
					<span data-group={group().id}>
						<Index each={group().controls}>{(control) => <Control control={control()} />}</Index>
					</span>
				)}
			</Index>
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
export function ShapesPane(props: { controller: ViewerController; locale?: string | undefined }) {
	let host!: HTMLDivElement;
	createEffect(() => {
		const stop = buildingBlock('mountDiagramShapes')(host, props.controller, {
			locale: props.locale,
		});
		onCleanup(stop);
	});
	return <div ref={host} />;
}
export type {
	BuildingBlockOptions,
	BuildingBlockHostCallbacks,
	DiagramToolbarProps,
} from '../../bindings/src/building-blocks';
export type { ToolbarGroupId } from '../../bindings/src/toolbar-controls';
