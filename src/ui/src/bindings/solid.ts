// `ooxml-ui/bindings/solid`: one Solid component per product, written without JSX so it needs no
// compiler: the component returns its host <div>, re-applies options reactively and hands the
// handle to a ref callback prop (`editorRef`). On the server it renders the empty host, which the
// browser claims while hydrating.
import { createEffect, createSignal, onCleanup, onMount } from 'solid-js';
import { getNextElement, isServer, ssrElement } from 'solid-js/web';
import type { MountedBinding } from './contract';
import { deferredHandle, type DeferredHandleOptions, type HandleSpec } from './handle';

export interface SolidBindingConfig<Props, O, B, H> {
	handle: HandleSpec<H>;
	handleOptions?: DeferredHandleOptions;
	mount(host: HTMLDivElement, options: O): B;
	/** The mount options; reading each prop here is what makes the effect track it. */
	options(props: Props): O;
	/** The callback prop that receives the handle once mounted. Default `editorRef`. */
	ref?: string;
}

export function createSolidBinding<
	Props extends { class?: string | undefined },
	H,
	O,
	B extends MountedBinding<O> = MountedBinding<O>,
>(config: SolidBindingConfig<Props, O, B, H>): (props: Props) => HTMLDivElement {
	const refName = config.ref ?? 'editorRef';
	return function Binding(props: Props): HTMLDivElement {
		if (isServer)
			return ssrElement(
				'div',
				{ class: props.class },
				undefined,
				true,
			) as unknown as HTMLDivElement;
		const host = getNextElement(() => document.createElement('div')) as HTMLDivElement;
		const [binding, setBinding] = createSignal<B>();
		onMount(() => {
			const mounted = config.mount(host, config.options(props));
			setBinding(() => mounted);
			const ref = (props as Record<string, unknown>)[refName];
			if (typeof ref === 'function')
				ref(deferredHandle(config.handle, () => mounted as unknown as H, config.handleOptions));
		});
		createEffect(() => {
			const current = binding();
			if (!current) return;
			host.className = props.class ?? '';
			current.update(config.options(props));
		});
		onCleanup(() => binding()?.destroy());
		return host;
	};
}
