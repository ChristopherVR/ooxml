import { defineComponent, h, onBeforeUnmount, onMounted, ref, watch, type PropType } from 'vue';
import {
	eventKeys,
	propertyKeys,
	mountFrameworkViewer,
	viewerHandle,
	withEventEmitter,
	type MountedViewer,
	type ViewerCallbacks,
	type ViewerProperties,
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
export type {
	ViewerHandle,
	ViewerCallbacks,
	ViewerOptions,
	ViewerEvents,
	ViewerEditState,
	VsdxExportResult,
} from './common.js';
