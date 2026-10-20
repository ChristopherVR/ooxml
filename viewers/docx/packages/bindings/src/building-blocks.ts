import { mountEditorSurface, type EditorHandle, type SurfaceOptions } from './index';
import type { WordToolbarProps } from 'docx-web-component';
export interface BuildingBlockOptions extends SurfaceOptions {
	bytes?: Uint8Array | ArrayBuffer | undefined;
}
export interface CanvasProps {
	mount: (host: HTMLElement | null) => void;
	className?: string | undefined;
}
export interface BuildingBlockState {
	toolbarProps: WordToolbarProps | null;
	loading: boolean;
	error: Error | null;
	mode: 'view' | 'edit';
}
/** Framework-neutral engine lifetime. Each adapter owns its native reactive state. */
export function createBuildingBlockSession(initial: BuildingBlockOptions = {}) {
	let options = initial;
	let binding: ReturnType<typeof mountEditorSurface> | undefined;
	let unsubscribe: (() => void) | undefined;
	let lastBytes: BuildingBlockOptions['bytes'];
	let generation = 0;
	let currentHost: HTMLElement | null = null;
	let state: BuildingBlockState = { toolbarProps: null, loading: false, error: null, mode: 'edit' };
	const listeners = new Set<(state: BuildingBlockState) => void>();
	const publish = (patch: Partial<BuildingBlockState>) => {
		state = {
			...state,
			...patch,
			mode: options.canEdit === false || options.readOnly ? 'view' : 'edit',
		};
		for (const listener of listeners) listener(state);
	};
	const refresh = () =>
		publish({
			toolbarProps: binding?.getToolbarProps() ?? null,
			loading: binding?.element.loading ?? false,
			...(binding?.element.loading ? { error: null } : {}),
		});
	const effective = (): SurfaceOptions => ({
		...options,
		onDocumentError: (error) => {
			publish({ error });
			options.onDocumentError?.(error);
		},
	});
	const requireBinding = () => {
		if (!binding) throw new Error('Editor is not mounted');
		return binding;
	};
	const load = async (input: Uint8Array | ArrayBuffer) => {
		const mounted = requireBinding();
		const token = ++generation;
		publish({ loading: true, error: null });
		try {
			await mounted.load(input);
		} catch (cause) {
			if (token === generation && binding === mounted)
				publish({ error: cause instanceof Error ? cause : new Error(String(cause)) });
			throw cause;
		} finally {
			if (token === generation && binding === mounted) refresh();
		}
	};
	const loadOptions = () => {
		if (!binding || !options.bytes || lastBytes === options.bytes) return;
		lastBytes = options.bytes;
		void load(options.bytes).catch(() => {
			/* The state and document error callback report it. */
		});
	};
	const mount = (host: HTMLElement | null) => {
		if (host === currentHost) return;
		currentHost = host;
		++generation;
		unsubscribe?.();
		unsubscribe = undefined;
		binding?.destroy();
		binding = undefined;
		lastBytes = undefined;
		if (host) {
			binding = mountEditorSurface(host, effective());
			unsubscribe = binding.subscribe(refresh);
		}
		refresh();
		loadOptions();
	};
	const handle: EditorHandle = {
		get element() {
			return requireBinding().element;
		},
		load,
		save: () => requireBinding().save(),
		download: (name) => requireBinding().download(name),
		markClean: () => requireBinding().markClean(),
		get dirty() {
			return binding?.dirty ?? false;
		},
	};
	return {
		canvasProps: { mount } satisfies CanvasProps,
		handle,
		get state() {
			return state;
		},
		subscribe(listener: (state: BuildingBlockState) => void) {
			listeners.add(listener);
			listener(state);
			return () => {
				listeners.delete(listener);
			};
		},
		update(next: BuildingBlockOptions) {
			if (options.documentModel !== next.documentModel) publish({ error: null });
			options = next;
			binding?.update(effective());
			refresh();
			loadOptions();
		},
		destroy() {
			mount(null);
			listeners.clear();
		},
	};
}
export type { WordToolbarProps } from 'docx-web-component';
