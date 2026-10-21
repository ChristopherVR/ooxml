import type { VisioDocument, VisioPage } from 'ooxml-core/visio';
import {
	visioMarqueeBox,
	visioMarqueeSelection,
	visioSelectionIsOnPage,
	type VisioPagePoint,
} from 'ooxml-core/visio/ui';
import { documentVisibility } from './viewer-layers';
import type { ViewerController, ViewerState } from './controller';
import type { createMovementPreview } from './viewer-movement-preview';

export interface Gesture {
	pointer: number;
	svg: SVGSVGElement;
	document: VisioDocument;
	page: VisioPage;
	selection: ViewerState['selectedShapes'];
	layers: ViewerState['layerVisibilityOverrides'];
	zoom: number;
	start: VisioPagePoint;
	clientX: number;
	clientY: number;
	started: boolean;
	additive: boolean;
	target?: { id: string; name: string; pageId: string };
	preview?: ReturnType<typeof createMovementPreview>;
	marquee?: SVGRectElement;
	showOverlays?: () => void;
	touch: boolean;
	/** The shape was found within the finger's hit tolerance, not under it. */
	nearby: boolean;
	longPress?: number | undefined;
	/** Set once the finger has been held still; lifting it then opens the context menu. */
	held?: (() => void) | undefined;
}

/** The gesture still matches the document, page, selection, layers, zoom and editing state. */
export function gestureIsCurrent(drag: Gesture, state: ViewerState, active: boolean): boolean {
	return (
		state.document === drag.document &&
		state.document?.pages[state.pageIndex] === drag.page &&
		state.selectedShapes === drag.selection &&
		state.layerVisibilityOverrides === drag.layers &&
		state.zoom === drag.zoom &&
		drag.svg.isConnected &&
		!state.loading &&
		!state.edit.busy &&
		(!drag.target || state.canEdit) &&
		active
	);
}
/**
 * Selects the dragged shape (as Visio does when an unselected shape is dragged) and reports
 * whether the resulting state is still the one the gesture intended.
 */
export function selectDragTarget(
	controller: ViewerController,
	drag: Gesture,
	active: boolean,
): boolean {
	const target = drag.target!;
	const intended = drag.additive ? [...drag.selection, target] : [target];
	controller.selectShapes(intended);
	const state = controller.state;
	if (
		state.document !== drag.document ||
		state.document.pages[state.pageIndex] !== drag.page ||
		state.layerVisibilityOverrides !== drag.layers ||
		state.loading ||
		state.edit.busy ||
		!active ||
		state.selectedShapes.length !== intended.length ||
		!state.selectedShapes.every(
			(shape, index) =>
				shape.id === intended[index]!.id && visioSelectionIsOnPage(shape, drag.page.id),
		)
	)
		return false;
	drag.selection = state.selectedShapes;
	return true;
}
/** The dashed-free translucent rectangle a marquee drag draws on the paper. */
export function createMarqueeRect(svg: SVGSVGElement): SVGRectElement {
	const rect = svg.ownerDocument.createElementNS('http://www.w3.org/2000/svg', 'rect');
	rect.dataset.marqueePreview = '';
	rect.setAttribute('fill', 'var(--_vv-accent)');
	rect.setAttribute('fill-opacity', '0.12');
	rect.setAttribute('stroke', 'var(--_vv-accent)');
	rect.setAttribute('stroke-width', '1.5');
	rect.setAttribute('vector-effect', 'non-scaling-stroke');
	rect.setAttribute('aria-hidden', 'true');
	rect.style.pointerEvents = 'none';
	svg.append(rect);
	return rect;
}
/** Builds the gesture for a pointerdown on `group` (a shape) or on empty paper (no group). */
export function createGesture(input: {
	event: PointerEvent;
	svg: SVGSVGElement;
	page: VisioPage;
	state: ViewerState;
	start: VisioPagePoint;
	touch: boolean;
	nearby: boolean;
	group: SVGGElement | null;
}): Gesture {
	const { event, state, group } = input;
	return {
		pointer: event.pointerId,
		svg: input.svg,
		page: input.page,
		document: state.document!,
		selection: state.selectedShapes,
		layers: state.layerVisibilityOverrides,
		zoom: state.zoom,
		start: input.start,
		clientX: event.clientX,
		clientY: event.clientY,
		started: false,
		additive: event.shiftKey || event.ctrlKey || event.metaKey,
		touch: input.touch,
		nearby: input.nearby,
		...(group
			? {
					target: {
						id: group.dataset.shapeId!,
						name: group.dataset.shapeName ?? '',
						pageId: input.page.id,
					},
				}
			: {}),
	};
}
/** Selects what a finished marquee fully encloses, adding to the selection when additive. */
export function marqueeSelect(
	controller: ViewerController,
	drag: Gesture,
	point: VisioPagePoint,
): void {
	const box = visioMarqueeBox(drag.start, point);
	if (!box) return;
	const selected = visioMarqueeSelection(
		drag.page,
		box,
		documentVisibility(drag.document, drag.layers),
	);
	controller.selectShapes(drag.additive ? [...drag.selection, ...selected] : selected);
}
