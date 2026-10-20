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
} from 'vue';
import {
	buildingBlock,
	createBuildingBlockSession,
	type BuildingBlockOptions,
	type DiagramToolbarProps,
	type ViewerController,
} from '../../bindings/src/building-blocks';
export function useViewerBuildingBlocks(options: MaybeRefOrGetter<BuildingBlockOptions> = {}) {
	const session = createBuildingBlockSession();
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
		return session.toolbar(toValue(options).locale);
	});
	return {
		controller: session.controller,
		state,
		canvasProps: { mount: session.mount },
		toolbarProps,
		inspectorProps: { controller: session.controller },
		loading: computed(() => state.value.loading),
		error: computed(() => state.value.error ?? state.value.edit.error),
		mode: computed(() => (state.value.canEdit ? ('edit' as const) : ('view' as const))),
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
export const Toolbar = defineComponent({
	name: 'DiagramToolbar',
	props: { toolbarProps: { type: Object as PropType<DiagramToolbarProps>, required: true } },
	setup(props) {
		return () =>
			h('div', { role: 'toolbar' }, [
				h(
					'button',
					{
						type: 'button',
						disabled: !props.toolbarProps.canUndo,
						onClick: props.toolbarProps.onUndo,
					},
					props.toolbarProps.labels.Undo,
				),
				h(
					'button',
					{
						type: 'button',
						disabled: !props.toolbarProps.canRedo,
						onClick: props.toolbarProps.onRedo,
					},
					props.toolbarProps.labels.Redo,
				),
				h(
					'button',
					{ type: 'button', onClick: props.toolbarProps.onFit },
					props.toolbarProps.labels['Fit to Window'],
				),
			]);
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
export type { BuildingBlockOptions, DiagramToolbarProps } from '../../bindings/src/building-blocks';
