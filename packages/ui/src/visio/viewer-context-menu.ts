import type { ViewerController } from './controller.js';
import { contextMenu } from './ribbon-parts.js';

const CLIPBOARD = 'Needs core shape copy and paste.';
const ARRANGE = 'Needs core grouping, container and z-order edits.';

/** Visio's shape and page context menus; commands the core lacks are shown disabled. */
export function createContextMenus(doc: Document): HTMLElement[] {
	return [
		contextMenu(doc, 'shape', 'Shape', [
			{ id: 'ctx-cut', label: 'Cut', icon: 'cut', unsupported: CLIPBOARD },
			{ id: 'ctx-copy', label: 'Copy', icon: 'copy', unsupported: CLIPBOARD },
			{ id: 'ctx-paste', label: 'Paste', icon: 'paste', unsupported: CLIPBOARD },
			'-',
			{
				id: 'ctx-edit-text',
				label: 'Edit Text',
				icon: 'pencil',
				action: { type: 'reveal', panel: 'edit', focusText: true },
			},
			'-',
			{ id: 'ctx-group', label: 'Group', icon: 'group', unsupported: ARRANGE },
			{ id: 'ctx-container', label: 'Container', icon: 'rectangle', unsupported: ARRANGE },
			{ id: 'ctx-order', label: 'Order', icon: 'bringToFront', unsupported: ARRANGE },
			'-',
			{
				id: 'ctx-hyperlink',
				label: 'Hyperlink...',
				icon: 'visioLink',
				unsupported: 'Needs core hyperlink edits.',
			},
			{
				id: 'ctx-shape-data',
				label: 'Shape Data',
				icon: 'visioShapeData',
				action: { type: 'reveal', panel: 'selection' },
			},
			{
				id: 'ctx-format',
				label: 'Format Shape',
				icon: 'fill',
				unsupported: 'Needs core fill, line and effect edits.',
			},
			'-',
			{
				id: 'ctx-comment',
				label: 'Add Comment',
				icon: 'message',
				unsupported: 'Needs core comments.',
			},
		]),
		contextMenu(doc, 'page', 'Page', [
			{ id: 'ctx-page-paste', label: 'Paste', icon: 'paste', unsupported: CLIPBOARD },
			'-',
			{
				id: 'ctx-fit',
				label: 'Fit to Window',
				icon: 'fitPage',
				action: { type: 'zoom', mode: 'fit' },
			},
			{ id: 'ctx-grid', label: 'Grid', icon: 'grid', action: { type: 'grid' } },
			{ id: 'ctx-ruler', label: 'Ruler', icon: 'ruler', action: { type: 'ruler' } },
			'-',
			{
				id: 'ctx-page-setup',
				label: 'Page Setup...',
				icon: 'visioPagesPane',
				unsupported: 'Needs core page setup edits.',
			},
			{
				id: 'ctx-page-comment',
				label: 'Add Comment',
				icon: 'message',
				unsupported: 'Needs core comments.',
			},
		]),
	];
}

type Menu = HTMLElement & { openAt(x: number, y: number): void };
const targetShape = (node: EventTarget | null) =>
	(node as Element | null)?.closest?.<SVGGElement>('[data-shape-id]');

/**
 * Right-click, Shift+F10 or the Menu key on the canvas: a shape menu for the shape under the
 * pointer (selecting it first, as Visio does) or a page menu on empty paper.
 */
export function wireContextMenus(
	root: ShadowRoot,
	viewport: HTMLElement,
	controller: ViewerController,
): () => void {
	const Abort = root.ownerDocument.defaultView?.AbortController ?? AbortController;
	const events = new Abort();
	const menu = (id: string) => root.querySelector<Menu>(`[data-context-menu="${id}"]`)!;
	const show = (shape: SVGGElement | null | undefined, x: number, y: number) => {
		const state = controller.state;
		if (!state.document?.pages[state.pageIndex]) return;
		if (shape) {
			controller.selectShape({
				id: shape.dataset.shapeId!,
				name: shape.dataset.shapeName ?? '',
				...(shape.dataset.pageId ? { pageId: shape.dataset.pageId } : {}),
			});
			// Selection may be refused (for example a hidden layer); fall back to the page menu.
			if (controller.state.selectedShape?.id === shape.dataset.shapeId)
				return menu('shape').openAt(x, y);
		}
		menu('page').openAt(x, y);
	};
	viewport.addEventListener(
		'contextmenu',
		(event) => {
			event.preventDefault();
			show(targetShape(event.target), event.clientX, event.clientY);
		},
		{ signal: events.signal },
	);
	viewport.addEventListener(
		'keydown',
		(event) => {
			if (!(event.key === 'ContextMenu' || (event.key === 'F10' && event.shiftKey))) return;
			event.preventDefault();
			const shape = targetShape(event.target);
			const box = (shape ?? viewport).getBoundingClientRect();
			show(shape, box.left + Math.min(24, box.width / 2), box.top + Math.min(24, box.height / 2));
		},
		{ signal: events.signal },
	);
	return () => events.abort();
}
