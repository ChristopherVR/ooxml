import * as officeTheme from 'ooxml-ui/theme';
import { mountEditorSurface, type EditorHandle, type SurfaceOptions } from './index';
import type { WordNavigationProps, WordStatusProps, WordToolbarProps } from 'docx-web-component';
export interface BuildingBlockOptions extends SurfaceOptions {
	bytes?: Uint8Array | ArrayBuffer | undefined;
}
export interface CanvasProps {
	mount: (host: HTMLElement | null) => void;
	className?: string | undefined;
}
export interface BuildingBlockState {
	toolbarProps: WordToolbarProps | null;
	/** Page, view mode and page navigation callbacks. Null until the canvas mounts. */
	navigationProps: WordNavigationProps | null;
	/** Word count and zoom. Null until the canvas mounts. */
	statusProps: WordStatusProps | null;
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
	let lastElement: unknown;
	let state: BuildingBlockState = {
		toolbarProps: null,
		navigationProps: null,
		statusProps: null,
		loading: false,
		error: null,
		mode: 'view',
	};
	const listeners = new Set<(state: BuildingBlockState) => void>();
	const publish = (patch: Partial<BuildingBlockState>) => {
		const next: BuildingBlockState = {
			...state,
			...patch,
			mode: options.canEdit !== true || options.readOnly ? 'view' : 'edit',
		};
		if (
			(Object.keys(next) as (keyof BuildingBlockState)[]).every((key) => next[key] === state[key])
		)
			return;
		state = next;
		for (const listener of listeners) listener(state);
	};
	// Snapshots of equal data keep their identity, so hosts re-render only on a real change. Their
	// callbacks read the live editor, so an older snapshot of the same engine stays valid.
	const stable = <T extends object>(previous: T | null, next: T | null, same: boolean): T | null =>
		previous && next && same && JSON.stringify(previous) === JSON.stringify(next) ? previous : next;
	const refresh = () => {
		const element = binding?.element ?? null;
		const same = lastElement === element;
		lastElement = element;
		publish({
			toolbarProps: stable(state.toolbarProps, binding?.getToolbarProps() ?? null, same),
			navigationProps: stable(state.navigationProps, binding?.getNavigationProps() ?? null, same),
			statusProps: stable(state.statusProps, binding?.getStatusProps() ?? null, same),
			loading: binding?.element.loading ?? false,
			...(binding?.element.loading ? { error: null } : {}),
		});
	};
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
export type { WordNavigationProps, WordStatusProps, WordToolbarProps } from 'docx-web-component';

/** Keep optional native controls on the shared Office styles without making old viewer imports fail. */
export function prepareWordToolbar(host: HTMLElement | null): void {
	if (!host) return;
	const prepare: unknown = Reflect.get(officeTheme, 'prepareOfficeUiBlock');
	if (typeof prepare === 'function') {
		(prepare as (host: HTMLElement, kind: 'toolbar', product: 'docx') => void)(
			host,
			'toolbar',
			'docx',
		);
	} else {
		officeTheme.installOfficeUiTheme(host.ownerDocument);
		host.setAttribute('data-office-block', 'toolbar');
		host.setAttribute('data-office-product', 'docx');
	}
}
