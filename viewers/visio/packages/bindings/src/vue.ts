import {
	defineComponent,
	h,
	onBeforeUnmount,
	onMounted,
	onScopeDispose,
	ref,
	shallowRef,
	watch,
	type PropType,
	type Ref,
	type ShallowRef,
} from 'vue';
import {
	eventKeys,
	propertyKeys,
	mountFrameworkViewer,
	viewerHandle,
	viewerStateSource,
	withEventEmitter,
	type MountedViewer,
	type ViewerCallbacks,
	type ViewerHandle,
	type ViewerProperties,
	type ViewerState,
} from './common.js';
/** Vue props and emitted event names are checked against the shared contract. */
export const VisioViewer = defineComponent({
	name: 'VisioViewer',
	props: {
		document: { type: Object as PropType<ViewerProperties['document']>, default: undefined },
		pageIndex: { type: Number, default: undefined },
		zoom: { type: Number, default: undefined },
		showToolbar: { type: Boolean, default: undefined },
		events: Object as PropType<ViewerCallbacks>,
	} satisfies Record<keyof ViewerProperties | 'events', unknown>,
	emits: [...eventKeys],
	setup(props, { emit, expose }) {
		const host = ref<HTMLElement>();
		let binding: MountedViewer | undefined;
		const options = () => withEventEmitter(props, (name, value) => emit(name, value));
		onMounted(() => {
			binding = mountFrameworkViewer(host.value!, options());
		});
		watch(
			() => [...propertyKeys.map((key) => props[key]), props.events],
			() => binding?.update(options()),
		);
		onBeforeUnmount(() => {
			const mounted = binding;
			binding = undefined;
			mounted?.destroy();
		});
		expose(viewerHandle(() => binding));
		return () => h('div', { ref: host });
	},
});
/**
 * Composable: reactive viewer state for a template ref to `<VisioViewer>`. The returned
 * shallow ref is `null` until the viewer mounts and follows remounts of the ref.
 */
export function useVisioViewerState(
	viewer: Ref<ViewerHandle | null | undefined>,
): Readonly<ShallowRef<ViewerState | null>> {
	const state = shallowRef<ViewerState | null>(null);
	let stop = () => {};
	watch(
		viewer,
		(handle) => {
			stop();
			const source = viewerStateSource(() => handle);
			state.value = source.getSnapshot();
			stop = source.subscribe(() => {
				state.value = source.getSnapshot();
			});
		},
		{ immediate: true, flush: 'post' },
	);
	onScopeDispose(() => stop());
	return state;
}
export type {
	ViewerState,
	ViewerHandle,
	ViewerCallbacks,
	ViewerOptions,
	ViewerEvents,
	ViewerEditState,
	VsdxExportResult,
} from './common.js';
