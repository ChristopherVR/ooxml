// `ooxml-ui/bindings/react`: one React component per product, generated from its mount function.
// React owns the host <div> and the lifecycle; the element owns all document behaviour.
import {
	createElement,
	forwardRef,
	useEffect,
	useImperativeHandle,
	useRef,
	type ForwardRefExoticComponent,
	type PropsWithoutRef,
	type RefAttributes,
} from 'react';
import type { MountedBinding } from './contract';
import { deferredHandle, type DeferredHandleOptions, type HandleSpec } from './handle';

export interface ReactBindingConfig<Props, O, B, H> {
	displayName: string;
	/** The members the component's `ref` exposes. */
	handle: HandleSpec<H>;
	handleOptions?: DeferredHandleOptions;
	/** Creates the element inside `host`; the binding it returns also serves the handle. */
	mount(host: HTMLDivElement, options: O): B;
	/** The mount options for a render's props; default the props themselves. */
	options?: (props: Props) => O;
	/** Attributes of the host `<div>`; default `{ className: props.className }`. */
	host?: (props: Props) => Record<string, unknown>;
}

const defaultHost = (props: object) => ({
	className: (props as { className?: string }).className,
});

/**
 * A `forwardRef` component that mounts once, forwards every render's props through
 * `binding.update` (the binding itself skips unchanged values) and exposes a handle that throws a
 * clear error before mount.
 */
export function createReactBinding<
	Props extends object,
	H,
	O = Props,
	B extends MountedBinding<O> = MountedBinding<O>,
>(
	config: ReactBindingConfig<Props, O, B, H>,
): ForwardRefExoticComponent<PropsWithoutRef<Props> & RefAttributes<H>> {
	const toOptions = config.options ?? ((props: Props) => props as unknown as O);
	const hostProps = config.host ?? defaultHost;
	const component = forwardRef<H, Props>(function Binding(props, ref) {
		const host = useRef<HTMLDivElement>(null);
		const binding = useRef<B | null>(null);
		useEffect(() => {
			const mounted = config.mount(host.current!, toOptions(props as Props));
			binding.current = mounted;
			return () => {
				binding.current = null;
				mounted.destroy();
			};
		}, []);
		useEffect(() => {
			binding.current?.update(toOptions(props as Props));
		}, [props]);
		useImperativeHandle(
			ref,
			() =>
				deferredHandle(
					config.handle,
					() => binding.current as unknown as H | null,
					config.handleOptions,
				),
			[],
		);
		return createElement('div', { ref: host, ...hostProps(props as Props) });
	});
	component.displayName = config.displayName;
	return component;
}
