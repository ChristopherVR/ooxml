import {
	createElement,
	useCallback,
	useEffect,
	useImperativeHandle,
	useRef,
	useState,
	type Ref,
} from 'react';
import {
	mountEditor,
	type EditorOptions,
	type EditorBinding,
	type EditorHandle,
} from '../../bindings/src/index';
import { prepareWordToolbar } from '../../bindings/src/building-blocks';
import type { WordToolbarProps } from 'docx-web-component';
export interface BuildingBlockOptions extends EditorOptions {
	canEdit?: boolean | undefined;
	handle?: Ref<EditorHandle> | undefined;
	bytes?: Uint8Array | ArrayBuffer | undefined;
}
export interface CanvasProps {
	mount: (host: HTMLDivElement | null) => void;
	className?: string | undefined;
}
/** A document surface with no application controls. Keep the returned props together. */
export function DocumentCanvas(props: CanvasProps) {
	return createElement('div', { ref: props.mount, className: props.className });
}
export function useViewerBuildingBlocks(options: BuildingBlockOptions = {}) {
	const latest = useRef(options);
	latest.current = options;
	const binding = useRef<EditorBinding | null>(null);
	const stop = useRef<(() => void) | undefined>(undefined);
	const [toolbarProps, setToolbarProps] = useState<WordToolbarProps | null>(null);
	const [loading, setLoading] = useState(false);
	const [mountVersion, setMountVersion] = useState(0);
	const [error, setError] = useState<Error | null>(null);
	const lastBytes = useRef<Uint8Array | ArrayBuffer | undefined>(undefined);
	const lastElement = useRef<EditorBinding['element'] | null>(null);
	const effective = () => ({
		...latest.current,
		readOnly: latest.current.canEdit === false || latest.current.readOnly === true,
		onDocumentError: (failure: Error) => {
			setError(failure);
			latest.current.onDocumentError?.(failure);
		},
	});
	const refresh = () => {
		const element = binding.current?.element ?? null;
		const sameElement = lastElement.current === element;
		lastElement.current = element;
		setLoading(binding.current?.element.loading ?? false);
		if (binding.current?.element.loading) setError(null);
		const next = binding.current?.element.getToolbarProps() ?? null;
		setToolbarProps((previous) =>
			sameElement && JSON.stringify(previous) === JSON.stringify(next) ? previous : next,
		);
	};
	const mount = useCallback((host: HTMLDivElement | null) => {
		stop.current?.();
		stop.current = undefined;
		binding.current?.destroy();
		binding.current = null;
		if (!host) {
			setToolbarProps(null);
			setLoading(false);
			return;
		}
		setMountVersion((version) => version + 1);
		lastBytes.current = undefined;
		binding.current = mountEditor(host, effective(), true);
		stop.current = binding.current.element.subscribeEditorState(refresh);
		refresh();
	}, []);
	useEffect(() => {
		setError(null);
	}, [options.documentModel]);
	useEffect(() => {
		binding.current?.update(effective());
	}, [options]);
	useEffect(() => {
		const bytes = options.bytes;
		if (!bytes || !binding.current || lastBytes.current === bytes) return;
		lastBytes.current = bytes;
		const mounted = binding.current;
		const active = () => binding.current === mounted && lastBytes.current === bytes;
		setLoading(true);
		setError(null);
		void binding.current
			.load(bytes)
			.catch((cause: unknown) => {
				if (active()) setError(cause instanceof Error ? cause : new Error(String(cause)));
			})
			.finally(() => {
				if (active()) {
					setLoading(false);
					refresh();
				}
			});
	}, [options.bytes, mountVersion]);
	const get = () => {
		if (!binding.current) throw new Error('Editor is not mounted');
		return binding.current;
	};
	useImperativeHandle(
		options.handle,
		() => ({
			get element() {
				return get().element;
			},
			load: async (input: Uint8Array | ArrayBuffer) => get().load(input),
			save: async () => get().save(),
			download: async (name?: string) => get().download(name),
			markClean: () => get().markClean(),
			get dirty() {
				return binding.current?.dirty ?? false;
			},
		}),
		[],
	);
	return {
		canvasProps: { mount } satisfies CanvasProps,
		toolbarProps,
		loading,
		error,
		mode: options.canEdit === false || options.readOnly ? ('view' as const) : ('edit' as const),
	};
}
export type { WordToolbarProps } from 'docx-web-component';

/** Optional controls using exactly the callbacks exposed to custom host controls. */
export function Toolbar(props: WordToolbarProps) {
	const buttons = [
		['Bold', props.bold, props.onBold, !props.canEdit],
		['Italic', props.italic, props.onItalic, !props.canEdit],
		['Underline', props.underline, props.onUnderline, !props.canEdit],
		['Undo', false, props.onUndo, !props.canUndo],
		['Redo', false, props.onRedo, !props.canRedo],
	] as const;
	return createElement(
		'div',
		{ role: 'toolbar', ref: prepareWordToolbar },
		...buttons.map(([label, pressed, run, disabled]) =>
			createElement(
				'button',
				{
					key: label,
					type: 'button',
					disabled,
					'aria-pressed': pressed === null ? 'mixed' : pressed,
					onMouseDown: (event: import('react').MouseEvent) => event.preventDefault(),
					onClick: () => {
						void run();
					},
				},
				props.labels[label] ?? label,
			),
		),
	);
}
