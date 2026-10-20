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
	ViewerController,
	createWorkerParser,
	type DiagramToolbarProps,
	type VsdxSource,
} from 'ooxml-ui/visio';

import * as diagramUi from 'ooxml-ui/visio';
type BuildingBlockApi = Pick<
	typeof diagramUi,
	'mountDiagramCanvas' | 'fitDiagramCanvas' | 'mountDiagramInspector' | 'diagramToolbarProps'
>;
/** Existing viewer imports also work against the previous runtime; new surfaces require its new API. */
function buildingBlock<K extends keyof BuildingBlockApi>(key: K): BuildingBlockApi[K] {
	const value: unknown = Reflect.get(diagramUi, key);
	if (typeof value !== 'function') throw new Error('Update ooxml-ui to use Visio building blocks.');
	return value as BuildingBlockApi[K];
}

export interface BuildingBlockOptions {
	content?: VsdxSource | undefined;
	canEdit?: boolean | undefined;
	locale?: string | undefined;
	handle?: Ref<ViewerController> | undefined;
	onDocumentChange?:
		| ((document: NonNullable<ViewerController['state']['document']>) => void)
		| undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
	onDocumentError?: ((error: Error) => void) | undefined;
}
export interface DiagramCanvasProps {
	mount(host: HTMLDivElement | null): void;
	className?: string | undefined;
}
export function DiagramCanvas(props: DiagramCanvasProps) {
	return createElement('div', { ref: props.mount, className: props.className });
}
/** Uses the same controller and SVG renderer as VisioViewer. The handle owns serialization. */
export function useViewerBuildingBlocks(options: BuildingBlockOptions = {}) {
	const [controller] = useState(() => new ViewerController(createWorkerParser()));
	const [state, setState] = useState(controller.state);
	const latest = useRef(options);
	latest.current = options;
	const host = useRef<HTMLDivElement | null>(null);
	const dispose = useRef<(() => void) | undefined>(undefined);
	const epoch = useRef(0);
	const mount = useCallback(
		(node: HTMLDivElement | null) => {
			dispose.current?.();
			dispose.current = undefined;
			host.current = node;
			if (node) dispose.current = buildingBlock('mountDiagramCanvas')(node, controller);
		},
		[controller],
	);
	useEffect(() => {
		const generation = ++epoch.current;
		const stop = controller.subscribe(setState);
		const events = controller.onEvent((name, detail) => {
			if (name === 'document-change') {
				const change = detail as { document: NonNullable<typeof state.document>; dirty: boolean };
				latest.current.onDocumentChange?.(change.document);
				latest.current.onDirtyChange?.(change.dirty);
			} else if (name === 'document-error') latest.current.onDocumentError?.(detail as Error);
		});
		return () => {
			stop();
			events();
			queueMicrotask(() => {
				if (epoch.current === generation) controller.destroy();
			});
		};
	}, [controller]);
	useEffect(() => {
		controller.canEdit = options.canEdit !== false;
	}, [controller, options.canEdit]);
	useEffect(() => {
		if (!options.content) return;
		let active = true;
		void (async () => {
			const bytes =
				options.content instanceof Blob
					? new Uint8Array(await options.content.arrayBuffer())
					: options.content;
			if (active && bytes) await controller.load(bytes);
		})().catch((error: unknown) => {
			if (active)
				latest.current.onDocumentError?.(error instanceof Error ? error : new Error(String(error)));
		});
		return () => {
			active = false;
			controller.cancelLoad();
		};
	}, [controller, options.content]);
	useImperativeHandle(options.handle, () => controller, [controller]);
	const fit = () => {
		if (host.current) buildingBlock('fitDiagramCanvas')(host.current, controller);
	};
	return {
		canvasProps: { mount } satisfies DiagramCanvasProps,
		toolbarProps: buildingBlock('diagramToolbarProps')(controller, fit, options.locale),
		inspectorProps: { controller },
		loading: state.loading,
		error: state.error ?? state.edit.error,
		mode: state.canEdit ? ('edit' as const) : ('view' as const),
	};
}

export function Toolbar(props: DiagramToolbarProps) {
	return createElement(
		'div',
		{ role: 'toolbar' },
		createElement(
			'button',
			{
				type: 'button',
				disabled: !props.canUndo,
				onClick: () => {
					void props.onUndo();
				},
			},
			props.labels.Undo,
		),
		createElement(
			'button',
			{
				type: 'button',
				disabled: !props.canRedo,
				onClick: () => {
					void props.onRedo();
				},
			},
			props.labels.Redo,
		),
		createElement(
			'button',
			{ type: 'button', onClick: props.onFit },
			props.labels['Fit to Window'],
		),
	);
}
export function PageNavigation(props: DiagramToolbarProps) {
	return createElement(
		'select',
		{
			'aria-label': props.labels.Page,
			value: props.pageIndex,
			onChange: (event: import('react').ChangeEvent<HTMLSelectElement>) =>
				props.onPageChange(Number(event.currentTarget.value)),
		},
		...props.pages.map((page, index) =>
			createElement('option', { key: page.id, value: index }, page.name),
		),
	);
}
export function Inspector(props: { controller: ViewerController }) {
	const host = useRef<HTMLDivElement>(null);
	useEffect(
		() =>
			host.current
				? buildingBlock('mountDiagramInspector')(host.current, props.controller)
				: undefined,
		[props.controller],
	);
	return createElement('div', { ref: host });
}
export type { DiagramToolbarProps } from 'ooxml-ui/visio';
