import { defineRuler, type OfficeUiRuler, type RulerMarkerEventDetail } from '../chrome/ruler';
import type { ParagraphFormat } from './paragraph-format';

/** Everything the ruler draws, in CSS pixels at 100% zoom (96 per inch). */
export interface RulerGeometry {
	pageWidth: number;
	marginLeft: number;
	marginRight: number;
	/** Current paragraph's indents, measured from the margins. */
	indentLeft: number;
	indentRight: number;
	/** Extra first-line offset from the left indent: positive first line, negative hanging. */
	firstLine: number;
}

export const PX_PER_INCH = 96;

/** Marker positions from the page's left edge: first-line, left indent, right indent. */
export function markerPositions(geometry: RulerGeometry): {
	firstLine: number;
	left: number;
	right: number;
} {
	const left = geometry.marginLeft + geometry.indentLeft;
	return {
		firstLine: left + geometry.firstLine,
		left,
		right: geometry.pageWidth - geometry.marginRight - geometry.indentRight,
	};
}

export type RulerMarker = 'first-line' | 'left' | 'right';

/** The paragraph fields a marker dragged to `x` (px from the page's left edge) changes. */
export function markerChange(
	marker: RulerMarker,
	x: number,
	geometry: RulerGeometry,
): Partial<Pick<ParagraphFormat, 'leftInches' | 'rightInches' | 'special' | 'specialInches'>> {
	const snap = (px: number) => Math.round(px / SNAP_PX) * SNAP_PX;
	const inches = (px: number) => Math.round((px / PX_PER_INCH) * 1000) / 1000;
	const positions = markerPositions(geometry);
	if (marker === 'right') {
		const edge = geometry.pageWidth - geometry.marginRight;
		const indent = Math.min(Math.max(snap(edge - x), 0), edge - positions.left - SNAP_PX);
		return { rightInches: inches(indent) };
	}
	if (marker === 'left') {
		// Like Word's left-indent marker, the first-line offset travels with it.
		const limit = geometry.pageWidth - geometry.marginRight - geometry.indentRight - SNAP_PX;
		const indent = Math.min(
			Math.max(snap(x - geometry.marginLeft), 0),
			limit - geometry.marginLeft,
		);
		return { leftInches: inches(indent) };
	}
	const offset = snap(x - positions.left);
	const first = Math.max(offset, -positions.left);
	return {
		special: first > 0 ? 'firstLine' : first < 0 ? 'hanging' : 'none',
		specialInches: inches(Math.abs(first)),
	};
}

const SNAP_PX = PX_PER_INCH / 16;
const geometries = new WeakMap<HTMLElement, RulerGeometry>();
type RulerChange = ReturnType<typeof markerChange>;

/** The geometry a marker dragged to `x` would produce, for the live preview. */
function previewGeometry(marker: RulerMarker, x: number, geometry: RulerGeometry): RulerGeometry {
	const preview = markerChange(marker, x, geometry);
	const shifted = { ...geometry };
	if (preview.leftInches !== undefined) shifted.indentLeft = preview.leftInches * PX_PER_INCH;
	if (preview.rightInches !== undefined) shifted.indentRight = preview.rightInches * PX_PER_INCH;
	if (preview.special)
		shifted.firstLine =
			(preview.special === 'hanging' ? -1 : 1) * (preview.specialInches ?? 0) * PX_PER_INCH;
	return shifted;
}

type RulerElement = HTMLElement &
	Pick<
		OfficeUiRuler,
		| 'extent'
		| 'marginStart'
		| 'marginEnd'
		| 'zoom'
		| 'origin'
		| 'label'
		| 'labelsWithinMargins'
		| 'markers'
	>;

function paintMarkers(ruler: RulerElement, geometry: RulerGeometry): void {
	const positions = markerPositions(geometry);
	ruler.markers = [
		{ name: 'first-line', position: positions.firstLine, edge: 'top' },
		{ name: 'left', position: positions.left },
		{ name: 'right', position: positions.right },
	];
}

/**
 * Word's horizontal ruler above the page, on the shared `office-ui-ruler`: inch ticks counted
 * from the left margin, the margins shaded, and draggable markers for the current paragraph's
 * first-line, left and right indents. Dragging snaps to sixteenths of an inch and applies as one
 * undoable step on release (`onChange`). Margin edges and tab stops are not draggable.
 */
export function createRuler(onChange?: (change: RulerChange) => void): HTMLElement {
	// Register the element here, like every other shared element the editor creates: without it the
	// ruler is an unknown inline element with no markers.
	defineRuler();
	const ruler = document.createElement('office-ui-ruler') as RulerElement;
	ruler.className = 'dve-ruler';
	ruler.label = 'Ruler';
	ruler.labelsWithinMargins = true;
	const detail = (event: Event) => (event as CustomEvent<RulerMarkerEventDetail>).detail;
	const marker = (event: Event) => detail(event).name as RulerMarker;
	ruler.addEventListener('ruler-marker-move', (event) => {
		const geometry = geometries.get(ruler);
		if (geometry)
			paintMarkers(ruler, previewGeometry(marker(event), detail(event).position, geometry));
	});
	ruler.addEventListener('ruler-marker-commit', (event) => {
		const geometry = geometries.get(ruler);
		if (geometry) onChange?.(markerChange(marker(event), detail(event).position, geometry));
	});
	ruler.addEventListener('ruler-marker-cancel', () => onChange?.({}));
	return ruler;
}

/** Positions the ruler's parts for `geometry`; `zoom` matches the page's CSS zoom. */
export function updateRuler(ruler: HTMLElement, geometry: RulerGeometry, zoom: number): void {
	const element = ruler as RulerElement;
	geometries.set(ruler, geometry);
	element.extent = geometry.pageWidth;
	element.zoom = zoom;
	element.origin = geometry.marginLeft;
	element.marginStart = geometry.marginLeft;
	element.marginEnd = geometry.marginRight;
	paintMarkers(element, geometry);
}
