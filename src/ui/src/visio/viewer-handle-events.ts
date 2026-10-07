import type { VisioDocument, VisioPage } from 'ooxml-core/visio';
import type { ViewerState } from './controller';
export interface HandleGestureSnapshot {
	pointer: number;
	svg: SVGSVGElement;
	page: VisioPage;
	document: VisioDocument;
	shapeId: string;
}
/** Cancel previews when source, page, selection, tool or editing availability changes. */
export function handleGestureIsCurrent(
	drag: HandleGestureSnapshot,
	state: ViewerState,
	active: boolean,
): boolean {
	return (
		state.document === drag.document &&
		state.document?.pages[state.pageIndex] === drag.page &&
		state.selectedShape?.id === drag.shapeId &&
		drag.svg.isConnected &&
		!state.loading &&
		!state.edit.busy &&
		active
	);
}
/** Shared capture, cancellation, Escape and post-release click ownership for geometry handles. */
export function wireHandleEvents(
	viewport: HTMLElement,
	callbacks: {
		pointer(): number | undefined;
		start(event: PointerEvent): void;
		move(event: PointerEvent): void;
		finish(event: PointerEvent): Promise<void>;
		cancel(): void;
	},
): () => void {
	const Abort = viewport.ownerDocument.defaultView?.AbortController ?? AbortController;
	const events = new Abort(),
		options = { signal: events.signal, capture: true };
	let suppressClick = false;
	viewport.addEventListener(
		'pointerdown',
		(event) => {
			suppressClick = false;
			callbacks.start(event);
		},
		options,
	);
	viewport.addEventListener('pointermove', (event) => callbacks.move(event), options);
	viewport.addEventListener(
		'pointerup',
		(event) => {
			if (event.pointerId === callbacks.pointer()) suppressClick = true;
			void callbacks.finish(event);
		},
		options,
	);
	const cancel = (event: PointerEvent) => {
		if (event.pointerId === callbacks.pointer()) callbacks.cancel();
	};
	viewport.addEventListener('pointercancel', cancel, options);
	viewport.addEventListener('lostpointercapture', cancel, options);
	viewport.addEventListener(
		'keydown',
		(event) => {
			if (event.key === 'Escape' && callbacks.pointer() !== undefined) {
				event.preventDefault();
				event.stopImmediatePropagation();
				suppressClick = true;
				callbacks.cancel();
			}
		},
		options,
	);
	viewport.addEventListener(
		'click',
		(event) => {
			if (suppressClick) {
				suppressClick = false;
				event.stopImmediatePropagation();
			}
		},
		options,
	);
	return () => {
		callbacks.cancel();
		events.abort();
	};
}
