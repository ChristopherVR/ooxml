import { mountViewer, type MountedViewer } from '../../../src/binding.js';
import {
	eventKeys,
	propertyKeys,
	type ViewerCallbacks,
	type ViewerOptions,
	type ViewerProperties,
} from '../../../src/contract.js';

/** Framework props permit undefined; omitted values preserve the shared binding's state. */
export type ViewerProps = { [K in keyof ViewerProperties]?: ViewerProperties[K] | undefined } & {
	events?: ViewerCallbacks | undefined;
};
export type ViewerHandle = Pick<MountedViewer, 'element' | 'controller' | 'load' | 'fit'>;

/** One mapping for every property and event, shared by every framework adapter. */
export function viewerOptions(props: ViewerProps): ViewerOptions {
	const result: ViewerOptions = { events: props.events ?? {} };
	for (const key of propertyKeys) {
		const value = props[key];
		if (value !== undefined) Object.assign(result, { [key]: value });
	}
	return result;
}

/**
 * Native props are snapshots, while the browser binding accepts explicit patches. Forward only
 * changed properties so a callback/class update cannot undo an imperative load or user zoom.
 */
export function mountFrameworkViewer(
	container: HTMLElement,
	initial: ViewerOptions,
): MountedViewer {
	const binding = mountViewer(container, initial);
	let previous: ViewerOptions = { ...initial };
	let destroyed = false;
	return {
		...binding,
		update(next) {
			if (destroyed) throw new Error('The viewer has been destroyed.');
			const patch: ViewerOptions = { events: next.events ?? {} };
			for (const key of propertyKeys) {
				if (next[key] !== undefined && next[key] !== previous[key])
					Object.assign(patch, { [key]: next[key] });
			}
			const before = previous;
			previous = { ...next };
			try {
				binding.update(patch);
			} catch (error) {
				previous = before;
				throw error;
			}
		},
		destroy() {
			if (destroyed) return;
			destroyed = true;
			previous = {};
			binding.destroy();
		},
	};
}

/** Lazy imperative methods reject both before mount and after unmount. */
export function viewerHandle(current: () => MountedViewer | undefined): ViewerHandle {
	const requireViewer = (): MountedViewer => {
		const binding = current();
		if (!binding) throw new Error('The viewer is not mounted.');
		return binding;
	};
	return {
		get element() {
			return requireViewer().element;
		},
		get controller() {
			return requireViewer().controller;
		},
		async load(source) {
			await requireViewer().load(source);
		},
		fit() {
			requireViewer().fit();
		},
	};
}

/** Native framework emits and callback-map props share one complete event inventory. */
export function withEventEmitter(
	props: ViewerProps,
	emit: (name: keyof ViewerCallbacks, value: unknown) => void,
): ViewerOptions {
	const options = viewerOptions(props);
	const callbacks: ViewerCallbacks = {};
	for (const name of eventKeys) {
		Object.assign(callbacks, {
			[name]: (value: unknown) => {
				try {
					(props.events?.[name] as ((value: unknown) => void) | undefined)?.(value);
				} finally {
					emit(name, value);
				}
			},
		});
	}
	return { ...options, events: callbacks };
}
export { mountViewer };
export type { MountedViewer, ViewerCallbacks, ViewerOptions, ViewerProperties };
export { eventKeys, propertyKeys };
export type { ViewerEvents, VsdxSource } from '../../../src/contract.js';
export { ViewerController, type ViewerState } from '../../../src/controller.js';
