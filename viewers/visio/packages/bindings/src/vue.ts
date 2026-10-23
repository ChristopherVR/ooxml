import { shallowRef, watchEffect, type PropType, type Ref, type ShallowRef } from 'vue';
import { defineVueBinding } from 'ooxml-ui/bindings/vue';
import {
	eventKeys,
	mountFrameworkViewer,
	VIEWER_CONTRACT,
	VIEWER_HANDLE_OPTIONS,
	viewerStateSource,
	withEventEmitter,
	type ViewerCallbacks,
	type ViewerHandle,
	type ViewerProperties,
	type ViewerProps,
	type ViewerState,
} from './common';
/** Vue props and emitted event names are checked against the shared contract. */
export const VisioViewer = defineVueBinding({
	name: 'VisioViewer',
	props: {
		canEdit: { type: Boolean, default: undefined },
		readOnly: { type: Boolean, default: undefined },
		document: { type: Object as PropType<ViewerProperties['document']>, default: undefined },
		pageIndex: { type: Number, default: undefined },
		zoom: { type: Number, default: undefined },
		showToolbar: { type: Boolean, default: undefined },
		ribbonAddIns: {
			type: Array as unknown as PropType<ViewerProperties['ribbonAddIns']>,
			default: undefined,
		},
		collaboration: {
			type: Object as PropType<ViewerProperties['collaboration']>,
			default: undefined,
		},
		storage: {
			type: Object as PropType<ViewerProperties['storage']>,
			default: undefined,
		},
		customization: {
			type: Object as PropType<ViewerProperties['customization']>,
			default: undefined,
		},
		events: Object as PropType<ViewerCallbacks>,
	} satisfies Record<keyof ViewerProperties | 'events', unknown>,
	contract: VIEWER_CONTRACT,
	// Events arrive through the shared callback map, so a host callback can still cancel one.
	emits: [...eventKeys],
	options: (props, emit) =>
		withEventEmitter(props as ViewerProps, (name, value) => emit(name, value)),
	handle: VIEWER_CONTRACT.handle,
	handleOptions: VIEWER_HANDLE_OPTIONS,
	mount: mountFrameworkViewer,
});
/**
 * Composable: reactive viewer state for a template ref to `<VisioViewer>`. The returned
 * shallow ref is `null` until the viewer mounts and follows remounts of the ref.
 */
export function useVisioViewerState(
	viewer: Ref<ViewerHandle | null | undefined>,
): Readonly<ShallowRef<ViewerState | null>> {
	const state = shallowRef<ViewerState | null>(null);
	watchEffect(
		(onCleanup) => {
			// A template ref can be exposed before onMounted creates its controller. Reading the
			// reactive binding through the handle follows that transition as well as ref changes.
			const source = viewerStateSource(() => viewer.value);
			state.value = source.getSnapshot();
			onCleanup(
				source.subscribe(() => {
					state.value = source.getSnapshot();
				}),
			);
		},
		{ flush: 'post' },
	);
	return state;
}
export type {
	ViewerState,
	ViewerHandle,
	ViewerCallbacks,
	ViewerOptions,
	ViewerEvents,
	VisioShapeSelection,
	CreateVsdxOptions,
	ViewerEditState,
	VsdxExportResult,
	RibbonAddInCommand,
	RibbonAddInGroup,
	RibbonAddInTab,
} from './common';
