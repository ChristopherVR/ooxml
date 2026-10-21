import { prepareBlock } from '../../bindings/src/block-theme';
import {
	computed,
	defineComponent,
	h,
	onBeforeUnmount,
	onScopeDispose,
	onMounted,
	shallowRef,
	watch,
	watchEffect,
	type PropType,
} from 'vue';
import {
	createBuildingBlockSession,
	emptyBuildingBlockState,
	toolbarModel,
	selectReference,
	type BuildingBlockOptions,
	type SpreadsheetToolbarProps,
} from '../../bindings/src/building-blocks';
export type {
	BuildingBlockOptions,
	SpreadsheetToolbarProps,
} from '../../bindings/src/building-blocks';
/** Call during setup. Pass a getter to track reactive input and permission changes. */
export function useViewerBuildingBlocks(options: () => BuildingBlockOptions = () => ({})) {
	const state = shallowRef(emptyBuildingBlockState());
	const session = createBuildingBlockSession((next) => {
		state.value = next;
	});
	watchEffect(() => session.update(options()));
	onScopeDispose(session.destroy);
	return {
		canvasProps: { mount: session.mount },
		handle: session.handle,
		toolbarProps: computed(() => state.value.toolbarProps),
		loading: computed(() => state.value.loading),
		error: computed(() => state.value.error),
		mode: computed(() => state.value.mode),
		dirty: computed(() => state.value.dirty),
		selection: computed(() => state.value.selection),
		collaboration: computed(() => state.value.collaboration),
		editError: computed(() => state.value.editError),
	};
}
export const SpreadsheetCanvas = defineComponent({
	name: 'SpreadsheetCanvas',
	props: {
		mount: { type: Function as PropType<(host: HTMLElement | null) => void>, required: true },
	},
	setup(props) {
		const host = shallowRef<HTMLElement>();
		let mounted: ((host: HTMLElement | null) => void) | undefined;
		const attach = () => {
			if (!host.value || mounted === props.mount) return;
			mounted?.(null);
			mounted = props.mount;
			mounted(host.value);
		};
		onMounted(attach);
		watch(() => props.mount, attach);
		onBeforeUnmount(() => {
			mounted?.(null);
			mounted = undefined;
		});
		return () => h('div', { ref: host });
	},
});
const stateProp = { type: Object as PropType<SpreadsheetToolbarProps>, required: true } as const;
export const Toolbar = defineComponent({
	name: 'SpreadsheetToolbar',
	props: { toolbarProps: stateProp },
	setup(props) {
		return () =>
			h(
				'div',
				{ role: 'toolbar', ref: (host) => prepareBlock(host, 'toolbar') },
				toolbarModel(props.toolbarProps).map((item) =>
					item.kind === 'select'
						? h(
								'select',
								{
									key: item.id,
									'aria-label': item.label,
									disabled: item.disabled,
									value: item.value,
									onChange: (e: Event) => {
										void item.change((e.target as HTMLSelectElement).value);
									},
								},
								item.options.map((o) =>
									h(
										'option',
										{ key: o.value, value: o.value, selected: o.value === item.value },
										o.label,
									),
								),
							)
						: h(
								'button',
								{
									key: item.id,
									type: 'button',
									disabled: item.disabled,
									'aria-pressed': item.pressed === null ? 'mixed' : item.pressed,
									onMousedown: (e: MouseEvent) => e.preventDefault(),
									onClick: () => {
										void item.run();
									},
								},
								item.label,
							),
				),
			);
	},
});
export const FormulaBar = defineComponent({
	name: 'SpreadsheetFormulaBar',
	props: { toolbarProps: stateProp },
	setup(props) {
		const invalid = shallowRef(false);
		const invalidName = shallowRef(false);
		const name = shallowRef(props.toolbarProps.selection);
		watch(
			() => props.toolbarProps.selection,
			(next) => {
				name.value = next;
				invalidName.value = false;
			},
		);
		return () =>
			h('div', { ref: (host) => prepareBlock(host, 'formula-bar') }, [
				h('input', {
					'aria-label': props.toolbarProps.labels['Name Box'],
					value: name.value,
					'aria-invalid': invalidName.value,
					onInput: (e: Event) => {
						name.value = (e.target as HTMLInputElement).value;
						invalidName.value = false;
					},
					onKeydown: (e: KeyboardEvent) => {
						if (e.key === 'Enter') {
							e.preventDefault();
							invalidName.value = !selectReference(props.toolbarProps, name.value);
						}
					},
				}),
				h('input', {
					'aria-label': props.toolbarProps.labels['Formula Bar'],
					value: props.toolbarProps.formula,
					readonly: !props.toolbarProps.canEdit,
					'aria-invalid': invalid.value || props.toolbarProps.editError !== null,
					onInput: (e: Event) => {
						const input = e.target as HTMLInputElement;
						invalid.value = false;
						props.toolbarProps.onFormulaInput(input.value, input.selectionEnd ?? undefined);
					},
					onKeydown: (e: KeyboardEvent) => {
						if (e.key === 'Enter') {
							e.preventDefault();
							invalid.value = !props.toolbarProps.onFormulaChange(
								(e.target as HTMLInputElement).value,
							);
						} else if (e.key === 'Escape') {
							e.preventDefault();
							props.toolbarProps.onCancelEdit();
							invalid.value = false;
						}
					},
				}),
				props.toolbarProps.editError === null
					? null
					: h('span', { role: 'alert' }, props.toolbarProps.editError),
			]);
	},
});
export const SheetTabs = defineComponent({
	name: 'SpreadsheetSheetTabs',
	props: { toolbarProps: stateProp },
	setup(props) {
		return () =>
			h(
				'div',
				{
					role: 'tablist',
					'aria-label': props.toolbarProps.labels['Sheet'],
					ref: (host) => prepareBlock(host, 'tabs'),
				},
				props.toolbarProps.sheets
					.filter((s) => s.state === 'visible')
					.map((s) =>
						h(
							'button',
							{
								type: 'button',
								role: 'tab',
								'aria-selected': props.toolbarProps.activeSheet === s.index,
								onClick: () => props.toolbarProps.onSheetChange(s.index),
							},
							s.name,
						),
					),
			);
	},
});
