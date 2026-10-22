import {
	computed,
	defineComponent,
	h,
	onScopeDispose,
	shallowRef,
	toValue,
	watchEffect,
	watch,
	type MaybeRefOrGetter,
	type PropType,
} from 'vue';
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
	type Control,
	type ControlGroup,
} from '../../bindings/src/controls';
export type {
	BuildingBlockOptions,
	CanvasProps,
	WordNavigationProps,
	WordStatusProps,
	WordToolbarProps,
} from '../../bindings/src/building-blocks';
export function useViewerBuildingBlocks(options: MaybeRefOrGetter<BuildingBlockOptions> = {}) {
	const session = createBuildingBlockSession();
	const state = shallowRef(session.state);
	const stop = session.subscribe((next) => {
		state.value = next;
	});
	watchEffect(() => session.update(toValue(options)));
	onScopeDispose(() => {
		stop();
		session.destroy();
	});
	return {
		canvasProps: session.canvasProps,
		handle: session.handle,
		toolbarProps: computed(() => state.value.toolbarProps),
		navigationProps: computed(() => state.value.navigationProps),
		statusProps: computed(() => state.value.statusProps),
		loading: computed(() => state.value.loading),
		error: computed(() => state.value.error),
		mode: computed(() => state.value.mode),
	};
}
export const DocumentCanvas = defineComponent({
	name: 'DocumentCanvas',
	inheritAttrs: false,
	props: { mount: { type: Function as PropType<CanvasProps['mount']>, required: true } },
	setup(props, { attrs }) {
		let host: HTMLElement | null = null;
		const attach = (node: unknown) => {
			const next = node as HTMLElement | null;
			if (host === next) return;
			host = next;
			props.mount(next);
		};
		watch(
			() => props.mount,
			(next, previous) => {
				previous(null);
				if (host) next(host);
			},
		);
		onScopeDispose(() => props.mount(null));
		return () => h('div', { ...attrs, ref: attach });
	},
});
function renderControl(control: Control) {
	if (control.kind === 'text') return h('span', { key: control.id, role: 'status' }, control.text);
	if (control.kind === 'select')
		return h(
			'select',
			{
				key: control.id,
				'aria-label': control.label,
				disabled: control.disabled,
				value: control.value,
				onChange: (event: Event) => control.change((event.target as HTMLSelectElement).value),
			},
			control.options.map((option) =>
				h('option', { key: option.value, value: option.value }, option.label),
			),
		);
	return h(
		'button',
		{
			key: control.id,
			type: 'button',
			'aria-label': control.label,
			disabled: control.disabled,
			...(control.kind === 'toggle' ? { 'aria-pressed': control.pressed } : {}),
			onMousedown: (event: MouseEvent) => event.preventDefault(),
			onClick: control.run,
		},
		control.label,
	);
}
/** Keyed nodes, so a control keeps its focus while the formatting state changes. */
function renderControls(role: string, groups: readonly ControlGroup[]) {
	return h(
		'div',
		{ role, ref: (host: unknown) => prepareWordToolbar(host as HTMLElement | null) },
		groups.map((group, index) =>
			h('span', { key: index, role: 'group' }, group.map(renderControl)),
		),
	);
}
/** Optional controls: history, font, formatting, alignment, lists and enabled host workflows. */
export const Toolbar = defineComponent({
	name: 'WordToolbar',
	props: { toolbarProps: { type: Object as PropType<WordToolbarProps>, required: true } },
	setup: (props) => () => renderControls('toolbar', toolbarControls(props.toolbarProps)),
});
/** Optional page navigation for `navigationProps`. Pages exist in Print Layout only. */
export const PageNavigation = defineComponent({
	name: 'WordPageNavigation',
	props: { navigationProps: { type: Object as PropType<WordNavigationProps>, required: true } },
	setup: (props) => () => renderControls('group', navigationControls(props.navigationProps)),
});
/** Optional word count and zoom for `statusProps`. */
export const StatusBar = defineComponent({
	name: 'WordStatusBar',
	props: { statusProps: { type: Object as PropType<WordStatusProps>, required: true } },
	setup: (props) => () => renderControls('group', statusControls(props.statusProps)),
});
