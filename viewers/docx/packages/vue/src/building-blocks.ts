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
	type BuildingBlockOptions,
	type CanvasProps,
	type WordToolbarProps,
} from '../../bindings/src/building-blocks';
export type {
	BuildingBlockOptions,
	CanvasProps,
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
export const Toolbar = defineComponent({
	name: 'WordToolbar',
	props: { toolbarProps: { type: Object as PropType<WordToolbarProps>, required: true } },
	setup(props) {
		return () =>
			h(
				'div',
				{ role: 'toolbar' },
				[
					['Bold', props.toolbarProps.bold, props.toolbarProps.onBold, !props.toolbarProps.canEdit],
					[
						'Italic',
						props.toolbarProps.italic,
						props.toolbarProps.onItalic,
						!props.toolbarProps.canEdit,
					],
					[
						'Underline',
						props.toolbarProps.underline,
						props.toolbarProps.onUnderline,
						!props.toolbarProps.canEdit,
					],
					['Undo', false, props.toolbarProps.onUndo, !props.toolbarProps.canUndo],
					['Redo', false, props.toolbarProps.onRedo, !props.toolbarProps.canRedo],
				].map(([label, pressed, run, disabled]) =>
					h(
						'button',
						{
							type: 'button',
							disabled,
							'aria-pressed': pressed === null ? 'mixed' : pressed,
							onMousedown: (event: MouseEvent) => event.preventDefault(),
							onClick: run,
						},
						props.toolbarProps.labels[String(label)] ?? String(label),
					),
				),
			);
	},
});
