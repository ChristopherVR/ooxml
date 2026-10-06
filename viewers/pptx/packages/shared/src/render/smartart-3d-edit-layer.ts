/**
 * The inline-editing layer over a 3D SmartArt view (framework-agnostic).
 *
 * `<pptx-three-view>` keeps a SmartArt's 2D SVG only as its fallback, hidden
 * once the scene is up, so a binding that edits node text on the canvas lays
 * a second render of the diagram over the scene with its SVG paint hidden:
 * its `[data-smartart-node-id]` groups still take the double-click and the
 * textarea it opens is not SVG, so it shows. That layer is an input surface,
 * not a second copy of the element, so its element markers go.
 *
 * @module smartart-3d-edit-layer
 */

/** Element markers a second render of a diagram must not repeat. */
export const EDIT_LAYER_MARKER_ATTRS = [
	'data-element-id',
	'data-testid',
	'data-pptx-element',
	'role',
	'aria-label',
	'aria-roledescription',
] as const;

/**
 * Controls the layer shows on top of the diagram copy (the node editor, the
 * fill swatches). They are real UI, so they keep their accessible name.
 */
const CONTROL_SELECTOR = 'button, textarea, input, select';

/** Remove the element markers from `root` and everything under it, except on its controls. */
export function stripEditLayerMarkers(root: Element): void {
	for (const el of [root, ...root.querySelectorAll('*')]) {
		if (el.matches(CONTROL_SELECTOR)) {
			continue;
		}
		for (const attr of EDIT_LAYER_MARKER_ATTRS) {
			if (el.hasAttribute(attr)) {
				el.removeAttribute(attr);
			}
		}
	}
}

/** Shapes a node group draws (the text and decoration carry no hit area of their own). */
const GEOMETRY_SELECTOR = 'path, rect, polygon, polyline, ellipse, circle';

/**
 * Whether viewport point `(x, y)` is on one of `node`'s drawn shapes, tested in
 * each shape's own coordinates so a turned or scaled diagram is hit where it is
 * drawn, not where its axis-aligned bounding box reaches. `null` when the node
 * has no shape this can be asked of (a host without SVG geometry).
 */
function nodeShapeContainsPoint(node: Element, x: number, y: number): boolean | null {
	let tested = false;
	for (const shape of node.querySelectorAll(GEOMETRY_SELECTOR)) {
		const geometry = shape as Partial<SVGGeometryElement>;
		const matrix = geometry.getScreenCTM?.();
		if (!matrix || typeof geometry.isPointInFill !== 'function') {
			continue;
		}
		const det = matrix.a * matrix.d - matrix.b * matrix.c;
		if (!det) {
			continue;
		}
		tested = true;
		const dx = x - matrix.e;
		const dy = y - matrix.f;
		const local = {
			x: (matrix.d * dx - matrix.c * dy) / det,
			y: (matrix.a * dy - matrix.b * dx) / det,
		};
		if (geometry.isPointInFill(local) || geometry.isPointInStroke?.(local)) {
			return true;
		}
	}
	return tested ? false : null;
}

/**
 * The `[data-smartart-node-id]` element under viewport point `(x, y)` inside
 * `root`, by geometry: the smallest node containing the point. Unlike
 * `document.elementsFromPoint`, this finds nodes that hit-testing skips, such
 * as a hidden layer's `pointer-events: none` groups.
 *
 * A node is tested against its drawn shapes where the host can, so a rotated
 * diagram resolves the node under the pointer rather than a neighbour whose
 * bounding box merely overlaps it; otherwise by its bounding box.
 */
export function smartArtNodeAtPoint(root: ParentNode, x: number, y: number): Element | null {
	let best: Element | null = null;
	let bestArea = Number.POSITIVE_INFINITY;
	for (const node of root.querySelectorAll('[data-smartart-node-id]')) {
		const box = node.getBoundingClientRect();
		const area = box.width * box.height;
		if (area <= 0 || area >= bestArea) {
			continue;
		}
		const onShape = nodeShapeContainsPoint(node, x, y);
		const hit = onShape ?? (x >= box.left && x <= box.right && y >= box.top && y <= box.bottom);
		if (hit) {
			best = node;
			bestArea = area;
		}
	}
	return best;
}

/** UI the layer hosts itself (editor, swatches): events on it are not routed to a node. */
const LAYER_UI_SELECTOR = 'textarea, input, button, select, [role="group"]';

/**
 * Route a 3D edit layer's pointer input to the node under the pointer, found by
 * geometry (`smartArtNodeAtPoint`) instead of by event target.
 *
 * The layer is an invisible copy under a perspective scene, and a node group is
 * only hit where it paints (often its label alone), so a double-click or hover
 * on the node's fill missed the group's own handlers. This forwards a
 * double-click as `dblclick` and a change of hovered node as `mouseover` to the
 * group, so each binding's existing editor and fill swatches run unchanged.
 * Returns a function that removes the listeners.
 */
export function routeEditLayerPointerToNodes(layer: Element): () => void {
	let forwarding = false;
	let hovered: Element | null = null;

	const forward = (node: Element, type: 'dblclick' | 'mouseover', source: MouseEvent): void => {
		forwarding = true;
		try {
			node.dispatchEvent(
				new MouseEvent(type, {
					bubbles: true,
					cancelable: true,
					clientX: source.clientX,
					clientY: source.clientY,
				}),
			);
		} finally {
			forwarding = false;
		}
	};
	// Only UI inside the layer counts: an ancestor of the layer (the selected
	// element's own `role="group"` wrapper, say) must not switch routing off.
	const onUi = (event: Event): boolean => {
		if (forwarding) {
			return true;
		}
		const ui = event.target instanceof Element ? event.target.closest(LAYER_UI_SELECTOR) : null;
		return ui !== null && ui !== layer && layer.contains(ui);
	};

	const onDblClick = (raw: Event): void => {
		const event = raw as MouseEvent;
		if (onUi(event)) {
			return;
		}
		const node = smartArtNodeAtPoint(layer, event.clientX, event.clientY);
		if (node) {
			event.stopPropagation();
			// Opening an editor hides the swatches; the next move over this node shows them again.
			hovered = null;
			forward(node, 'dblclick', event);
		}
	};
	const onMouseMove = (raw: Event): void => {
		const event = raw as MouseEvent;
		if (onUi(event)) {
			return;
		}
		const node = smartArtNodeAtPoint(layer, event.clientX, event.clientY);
		if (node && node !== hovered) {
			forward(node, 'mouseover', event);
		}
		hovered = node;
	};

	layer.addEventListener('dblclick', onDblClick, true);
	layer.addEventListener('mousemove', onMouseMove, true);
	return () => {
		layer.removeEventListener('dblclick', onDblClick, true);
		layer.removeEventListener('mousemove', onMouseMove, true);
	};
}
