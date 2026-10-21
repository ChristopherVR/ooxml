import { effectiveMode, prepareDiagramBlock } from '../../bindings/src/building-blocks';
import {
	toolbarGroups,
	type ToolbarControl,
	type ToolbarGroupId,
} from '../../bindings/src/toolbar-controls';
import {
	createElement,
	useCallback,
	useEffect,
	useImperativeHandle,
	useLayoutEffect,
	useRef,
	useState,
	type ReactElement,
	type Ref,
} from 'react';
import {
	ViewerController,
	createWorkerParser,
	type DiagramToolbarProps,
	type VsdxSource,
} from 'ooxml-ui/visio';
import type { BuildingBlockHostCallbacks } from '../../bindings/src/building-blocks';

import * as diagramUi from 'ooxml-ui/visio';
type BuildingBlockApi = Pick<
	typeof diagramUi,
	| 'mountDiagramCanvas'
	| 'fitDiagramCanvas'
	| 'mountDiagramInspector'
	| 'mountDiagramShapes'
	| 'diagramToolbarProps'
>;
/** Layout effects warn on the server; permission must still land before paint in the browser. */
const useIsomorphicLayoutEffect = typeof window === 'undefined' ? useEffect : useLayoutEffect;
/** Existing viewer imports also work against the previous runtime; new surfaces require its new API. */
function buildingBlock<K extends keyof BuildingBlockApi>(key: K): BuildingBlockApi[K] {
	const value: unknown = Reflect.get(diagramUi, key);
	if (typeof value !== 'function') throw new Error('Update ooxml-ui to use Visio building blocks.');
	return value as BuildingBlockApi[K];
}

const toolbarRef = (host: HTMLDivElement | null) => prepareDiagramBlock(host, 'toolbar');
const tabsRef = (host: HTMLSelectElement | null) => prepareDiagramBlock(host, 'tabs');

export interface BuildingBlockOptions extends BuildingBlockHostCallbacks {
	content?: VsdxSource | undefined;
	/** Host edit permission. Read-only until `true`, from the very first render. */
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
	const [controller] = useState(() => {
		const created = new ViewerController(createWorkerParser());
		// Permission is in place before the canvas mounts, so it is never briefly editable.
		created.canEdit = options.canEdit === true;
		return created;
	});
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
	useIsomorphicLayoutEffect(() => {
		controller.canEdit = options.canEdit === true;
	}, [controller, options.canEdit]);
	useEffect(() => {
		if (!options.content) return;
		const content = options.content;
		void controller
			.loadSource(async () =>
				content instanceof Blob ? new Uint8Array(await content.arrayBuffer()) : content,
			)
			.catch(() => {});
		return () => {
			controller.cancelLoad();
		};
	}, [controller, options.content]);
	useImperativeHandle(options.handle, () => controller, [controller]);
	const fit = () => {
		if (host.current) buildingBlock('fitDiagramCanvas')(host.current, controller);
	};
	return {
		canvasProps: { mount } satisfies DiagramCanvasProps,
		toolbarProps: buildingBlock('diagramToolbarProps')(controller, fit, options.locale, {
			onOpenFile: options.onOpenFile,
			onOpenPageSetup: options.onOpenPageSetup,
			onOpenSettings: options.onOpenSettings,
			onOpenShareDialog: options.onOpenShareDialog,
		}),
		inspectorProps: { controller },
		shapesProps: { controller, locale: options.locale } satisfies ShapesPaneProps,
		loading: state.loading,
		error: state.error ?? state.edit.error,
		/** Effective permission: host `canEdit` and a drawing that can be saved back. */
		mode: effectiveMode(state),
		/** Acknowledge a host save: clears dirty and fires `onDirtyChange(false)`. */
		markClean: () => controller.markClean(),
	};
}

function renderControl(control: ToolbarControl): ReactElement {
	if (control.kind === 'text')
		return createElement('span', { key: control.id, 'data-control': control.id }, control.text);
	if (control.kind === 'button')
		return createElement(
			'button',
			{
				key: control.id,
				type: 'button',
				'data-control': control.id,
				disabled: control.disabled,
				'aria-pressed': control.pressed,
				onClick: control.run,
			},
			control.label,
		);
	return createElement(
		'span',
		{ key: control.id, 'data-control': control.id },
		createElement('input', {
			type: 'search',
			'aria-label': control.label,
			placeholder: control.label,
			value: control.query,
			onChange: (event: import('react').ChangeEvent<HTMLInputElement>) =>
				control.change(event.currentTarget.value),
		}),
		createElement('span', { 'aria-live': 'polite' }, control.summary),
		renderControl(control.previous),
		renderControl(control.next),
	);
}
/**
 * Optional library toolbar: history, zoom and fit, selection editing, clipboard, drawing tools,
 * search and the host's dialog buttons. `include` picks groups; hosts may ignore it and build
 * their own controls from `toolbarProps`.
 */
export function Toolbar(props: DiagramToolbarProps & { include?: readonly ToolbarGroupId[] }) {
	return createElement(
		'div',
		{ role: 'toolbar', ref: toolbarRef },
		...toolbarGroups(props, props.include).map((group) =>
			createElement(
				'span',
				{ key: group.id, 'data-group': group.id },
				...group.controls.map(renderControl),
			),
		),
	);
}
export function PageNavigation(props: DiagramToolbarProps) {
	return createElement(
		'select',
		{
			'aria-label': props.labels.Page,
			ref: tabsRef,
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
export interface ShapesPaneProps {
	controller: ViewerController;
	locale?: string | undefined;
}
/** The optional Shapes (stencil) pane; masters it adds go through the mounted canvas's controller. */
export function ShapesPane(props: ShapesPaneProps) {
	const host = useRef<HTMLDivElement>(null);
	useEffect(
		() =>
			host.current
				? buildingBlock('mountDiagramShapes')(host.current, props.controller, {
						locale: props.locale,
					})
				: undefined,
		[props.controller, props.locale],
	);
	return createElement('div', { ref: host });
}
export type { DiagramToolbarProps } from 'ooxml-ui/visio';
export type { BuildingBlockHostCallbacks } from '../../bindings/src/building-blocks';
export type { ToolbarGroupId } from '../../bindings/src/toolbar-controls';
