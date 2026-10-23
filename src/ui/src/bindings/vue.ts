// `ooxml-ui/bindings/vue`: one Vue component per product, generated from its contract: the props
// it declares, kebab-case emits for every element event, `v-model` for modelled events
// (`v-model:readOnly`) and the handle on the template ref (`expose`).
import {
	defineComponent,
	h,
	onBeforeUnmount,
	onMounted,
	ref,
	shallowRef,
	watch,
	type ComponentObjectPropsOptions,
} from 'vue';
import { pickProps, type AnyContract, type MountedBinding } from './contract';
import { deferredHandle, type DeferredHandleOptions, type HandleSpec } from './handle';

type Emit = (event: string, ...args: unknown[]) => void;

export interface VueBindingConfig<PO, O, B, H> {
	name: string;
	/** Vue's runtime prop declarations (types, Boolean casting, defaults). */
	props: PO;
	/** Generates the emits and the default mount options. */
	contract: Pick<AnyContract, 'props' | 'events'>;
	/** Options that are not element events, emitted under their own name (`{ ready: 'onReady' }`). */
	callbacks?: Readonly<Record<string, string>>;
	/** Replaces the generated emits. */
	emits?: readonly string[];
	handle: HandleSpec<H>;
	/** Default `{ lenientGetters: true }`: a template may read `element` before mount. */
	handleOptions?: DeferredHandleOptions;
	mount(host: HTMLElement, options: O): B;
	/** Replaces the generated mount options (props plus one emitting handler per event). */
	options?: (props: Record<string, unknown>, emit: Emit) => O;
}

/** The emits a contract generates: every event, `update:<prop>` per model, then the callbacks. */
export function vueEmits(
	contract: Pick<AnyContract, 'events'>,
	callbacks: Readonly<Record<string, string>> = {},
): string[] {
	const models = Object.values(contract.events).flatMap((rule) =>
		rule.model ? [`update:${rule.model.prop}`] : [],
	);
	return [...Object.keys(contract.events), ...models, ...Object.keys(callbacks)];
}

/** Handler options that re-emit each event (and its model) from a Vue component. */
export function vueEmitters(
	contract: Pick<AnyContract, 'events'>,
	props: Record<string, unknown>,
	emit: Emit,
	callbacks: Readonly<Record<string, string>> = {},
): Record<string, (...args: unknown[]) => void> {
	const options: Record<string, (...args: unknown[]) => void> = {};
	for (const [name, rule] of Object.entries(contract.events)) {
		const model = rule.model;
		options[rule.option] = (...args) => {
			emit(name, ...args);
			if (!model) return;
			const value = model.value(...(args as never[]));
			if (value !== props[model.prop]) emit(`update:${model.prop}`, value);
		};
	}
	for (const [name, option] of Object.entries(callbacks))
		options[option] = (...args) => emit(name, ...args);
	return options;
}

export function defineVueBinding<
	PO extends ComponentObjectPropsOptions,
	H,
	O = Record<string, unknown>,
	B extends MountedBinding<O> = MountedBinding<O>,
>(config: VueBindingConfig<PO, O, B, H>) {
	const callbacks = config.callbacks ?? {};
	const keys = Object.keys(config.props);
	return defineComponent({
		name: config.name,
		props: config.props,
		emits: [...(config.emits ?? vueEmits(config.contract, callbacks))],
		setup(props, { emit, expose }) {
			const values = props as Record<string, unknown>;
			const host = ref<HTMLElement>();
			// Reactive, so state composables that read through the handle follow mount and unmount.
			const binding = shallowRef<B>();
			const options = (): O =>
				config.options
					? config.options(values, emit as Emit)
					: ({
							...pickProps(config.contract, values),
							...vueEmitters(config.contract, values, emit as Emit, callbacks),
						} as O);
			onMounted(() => {
				binding.value = config.mount(host.value!, options());
			});
			watch(
				() => keys.map((key) => values[key]),
				() => binding.value?.update(options()),
			);
			onBeforeUnmount(() => {
				const mounted = binding.value;
				binding.value = undefined;
				mounted?.destroy();
			});
			expose(
				deferredHandle(
					config.handle,
					() => binding.value as unknown as H | undefined,
					config.handleOptions ?? { lenientGetters: true },
				) as Record<string, unknown>,
			);
			return () => h('div', { ref: host });
		},
	});
}
