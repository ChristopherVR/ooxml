import {
	computed,
	defineComponent,
	h,
	onBeforeUnmount,
	onScopeDispose,
	shallowRef,
	toValue,
	watchEffect,
	watch,
	type MaybeRefOrGetter,
	type PropType,
	type VNode,
} from 'vue';
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
	type ToolbarControl,
	type ToolbarGroupId,
} from '../../bindings/src/toolbar-controls';
export function useViewerBuildingBlocks(options: MaybeRefOrGetter<BuildingBlockOptions> = {}) {
	// Permission is applied at creation, before the canvas can mount or be used.
	const session = createBuildingBlockSession(toValue(options));
	const state = shallowRef(session.controller.state);
	const stop = session.controller.subscribe((value) => {
		state.value = value;
	});
	watchEffect(() => session.update(toValue(options)));
	onScopeDispose(() => {
		stop();
		session.destroy();
	});
	const toolbarProps = computed(() => {
		void state.value;
		return session.toolbar(toValue(options));
	});
	return {
		controller: session.controller,
		state,
		canvasProps: { mount: session.mount },
		toolbarProps,
		inspectorProps: { controller: session.controller },
		shapesProps: computed(() => ({
			controller: session.controller,
			locale: toValue(options).locale,
		})),
		/** Acknowledge a host save: clears dirty and fires `onDirtyChange(false)`. */
		markClean: () => session.markClean(),
		loading: computed(() => state.value.loading),
		error: computed(() => state.value.error ?? state.value.edit.error),
		/** Effective permission: host `canEdit` and a drawing that can be saved back. */
		mode: computed(() => effectiveMode(state.value)),
	};
}
export const DiagramCanvas = defineComponent({
	name: 'DiagramCanvas',
	props: {
		mount: { type: Function as PropType<(node: HTMLElement | null) => void>, required: true },
	},
	setup(props) {
		let host: HTMLElement | null = null;
		let mounted = props.mount;
		const mount = (node: unknown) => {
			const next = node as HTMLElement | null;
			if (next === host) return;
			host = next;
			mounted(next);
		};
		watch(
			() => props.mount,
			(next) => {
				mounted(null);
				mounted = next;
				if (host) mounted(host);
			},
		);
		onBeforeUnmount(() => {
			mounted(null);
			host = null;
		});
		return () => h('div', { ref: mount });
	},
});
function renderControl(control: ToolbarControl): VNode {
	if (control.kind === 'text')
		return h('span', { key: control.id, 'data-control': control.id }, control.text);
	if (control.kind === 'button')
		return h(
			'button',
			{
				key: control.id,
				type: 'button',
				'data-control': control.id,
				disabled: control.disabled,
				'aria-pressed': control.pressed,
				onClick: control.run,
			},
			control.label,
		);
	return h('span', { key: control.id, 'data-control': control.id }, [
		h('input', {
			type: 'search',
			'aria-label': control.label,
			placeholder: control.label,
			value: control.query,
			onInput: (event: Event) => control.change((event.target as HTMLInputElement).value),
		}),
		h('span', { 'aria-live': 'polite' }, control.summary),
		renderControl(control.previous),
		renderControl(control.next),
	]);
}
/** Optional library toolbar; `include` picks groups. Hosts may build their own from toolbarProps. */
export const Toolbar = defineComponent({
	name: 'DiagramToolbar',
	props: {
		toolbarProps: { type: Object as PropType<DiagramToolbarProps>, required: true },
		include: { type: Array as PropType<readonly ToolbarGroupId[]>, required: false },
	},
	setup(props) {
		return () =>
			h(
				'div',
				{
					role: 'toolbar',
					ref: (host: unknown) => prepareDiagramBlock(host as HTMLElement | null, 'toolbar'),
				},
				toolbarGroups(props.toolbarProps, props.include).map((group) =>
					h('span', { key: group.id, 'data-group': group.id }, group.controls.map(renderControl)),
				),
			);
	},
});
export const PageNavigation = defineComponent({
	name: 'DiagramPageNavigation',
	props: { toolbarProps: { type: Object as PropType<DiagramToolbarProps>, required: true } },
	setup(props) {
		return () =>
			h(
				'select',
				{
					'aria-label': props.toolbarProps.labels.Page,
					ref: (host: unknown) => prepareDiagramBlock(host as HTMLElement | null, 'tabs'),
					value: props.toolbarProps.pageIndex,
					onChange: (event: Event) =>
						props.toolbarProps.onPageChange(Number((event.target as HTMLSelectElement).value)),
				},
				props.toolbarProps.pages.map((page, index) =>
					h('option', { value: index, key: page.id }, page.name),
				),
			);
	},
});
export const Inspector = defineComponent({
	name: 'DiagramInspector',
	props: { controller: { type: Object as PropType<ViewerController>, required: true } },
	setup(props) {
		const host = shallowRef<HTMLElement | null>(null);
		watchEffect((onCleanup) => {
			if (host.value)
				onCleanup(buildingBlock('mountDiagramInspector')(host.value, props.controller));
		});
		return () => h('div', { ref: host });
	},
});
export const ShapesPane = defineComponent({
	name: 'DiagramShapesPane',
	props: {
		controller: { type: Object as PropType<ViewerController>, required: true },
		locale: { type: String as PropType<string | undefined>, required: false },
	},
	setup(props) {
		const host = shallowRef<HTMLElement | null>(null);
		watchEffect((onCleanup) => {
			if (host.value)
				onCleanup(
					buildingBlock('mountDiagramShapes')(host.value, props.controller, {
						locale: props.locale,
					}),
				);
		});
		return () => h('div', { ref: host });
	},
});
export type {
	BuildingBlockOptions,
	BuildingBlockHostCallbacks,
	DiagramToolbarProps,
} from '../../bindings/src/building-blocks';
export type { ToolbarGroupId } from '../../bindings/src/toolbar-controls';
