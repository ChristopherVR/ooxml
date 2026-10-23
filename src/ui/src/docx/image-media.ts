import { DOMSerializer, type Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';
import { NodeSelection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import { createResizeHandle } from './resize-handle';
import { schema } from './schema';
import { chartNodeSpec, placementClass } from 'ooxml-core/docx/ui';
import { parseDiagram, smartArtNodeView } from './smartart-node-view';
import { shapeNodeView } from './shape-node-view';
import type { ThemeCatalog } from 'ooxml-core/docx';

/**
 * Resolves inline picture bytes (kept off the JSON model by docx-core) into object URLs. URLs are
 * cached per package part and revoked together when the owning editor releases the cache.
 */
export class ImageMediaCache {
	private readonly urls = new Map<string, string>();
	private readonly listeners = new Set<() => void>();
	constructor(private readonly lookup: (partName: string) => Uint8Array | undefined) {}

	urlFor(partName: string, contentType: string): string | undefined {
		const cached = this.urls.get(partName);
		if (cached) return cached;
		const bytes = this.lookup(partName);
		if (!bytes || typeof URL.createObjectURL !== 'function') return undefined;
		const url = URL.createObjectURL(new Blob([bytes.slice()], { type: contentType }));
		this.urls.set(partName, url);
		return url;
	}

	release(): void {
		if (typeof URL.revokeObjectURL === 'function')
			for (const url of this.urls.values()) URL.revokeObjectURL(url);
		this.urls.clear();
	}
	/** Refresh mounted images when their bytes arrive after the document reference. */
	refresh(): void {
		this.release();
		for (const listener of this.listeners) listener();
	}
	onChange(listener: () => void): () => void {
		this.listeners.add(listener);
		return () => {
			this.listeners.delete(listener);
		};
	}
}

export interface ImageNodeViewOptions {
	/** Opens Format Picture for the picture at `pos` (double-click). */
	editPicture?(pos: number): void;
	/** Opens Edit Data for the selected chart (double-click on a chart). */
	editChart?(): void;
	/** The widest a picture may be resized to, in CSS pixels. */
	maxWidth?(): number;
	/** The document theme, for resolving SmartArt and chart colours. */
	theme?(): ThemeCatalog | undefined;
}

const MIN_SIZE = 16;

/**
 * Sets a picture's or chart's size in the document, keeping every other attribute and the node
 * selected; one undo step. Saving writes the size to the drawing's `wp:extent` (and `a:ext`).
 */
export function resizePicture(view: EditorView, pos: number, widthPx: number, heightPx: number) {
	const node = view.state.doc.nodeAt(pos);
	if (!node || node.type !== schema.nodes.image) return;
	const selected =
		view.state.selection instanceof NodeSelection && view.state.selection.from === pos;
	const tr = view.state.tr.setNodeMarkup(pos, undefined, {
		...node.attrs,
		widthPx: Math.max(MIN_SIZE, Math.round(widthPx)),
		heightPx: Math.max(MIN_SIZE, Math.round(heightPx)),
	});
	if (selected) tr.setSelection(NodeSelection.create(tr.doc, pos));
	view.dispatch(closeHistory(tr));
}

/** The size, limits and commit every inline drawing's resize handle shares. */
function sizeOf(
	node: ProseMirrorNode,
	view: EditorView,
	getPos: () => number | undefined,
	options: ImageNodeViewOptions,
) {
	return {
		editable: () => view.editable,
		size: () => ({ width: Number(node.attrs.widthPx), height: Number(node.attrs.heightPx) }),
		min: MIN_SIZE,
		maxWidth: () => options.maxWidth?.() ?? Number.POSITIVE_INFINITY,
		commit: (width: number, height: number) => {
			const pos = getPos();
			if (pos !== undefined) resizePicture(view, pos, width, height);
		},
	};
}

/**
 * A chart: painted with the document theme inside a frame with Word's corner resize handle (free,
 * Shift keeps the aspect ratio). An edit or resize replaces the node, which repaints it at its
 * size. A right press selects it first, so the context menu offers the chart commands.
 */
function chartNodeView(
	node: ProseMirrorNode,
	view: EditorView,
	getPos: () => number | undefined,
	options: ImageNodeViewOptions,
): NodeView {
	const chart = DOMSerializer.renderSpec(document, chartNodeSpec(node, options.theme?.()))
		.dom as HTMLElement;
	const dom = document.createElement('span');
	dom.className = 'dve-chart-frame';
	dom.append(chart);
	const select = () => {
		const pos = getPos();
		if (pos === undefined) return false;
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
		return true;
	};
	chart.addEventListener('dblclick', () => {
		if (view.editable && select()) options.editChart?.();
	});
	chart.addEventListener('contextmenu', () => {
		const { selection } = view.state;
		if (!(selection instanceof NodeSelection && selection.from === getPos())) select();
	});
	const handle = createResizeHandle('dve-picture-handle', {
		...sizeOf(node, view, getPos, options),
		keepRatio: (event) => event.shiftKey,
		zoomOf: () => chart,
		preview: (width, height) => {
			chart.style.width = `${width}px`;
			chart.style.height = `${height}px`;
		},
	});
	dom.append(handle);
	return {
		dom,
		ignoreMutation: () => true,
		selectNode: () => dom.classList.add('dve-chart-selected'),
		deselectNode: () => dom.classList.remove('dve-chart-selected'),
		stopEvent: (event) => event.target === handle,
	};
}

function renderSpec(node: ProseMirrorNode): HTMLElement {
	// The placeholder spec carries child spans (a text box's lines), so render the whole spec.
	return DOMSerializer.renderSpec(document, schema.nodes.image.spec.toDOM!(node))
		.dom as HTMLElement;
}

/**
 * Renders pictures from package media and, for editable real pictures, adds Word's proportional
 * corner resize handle. The new size is applied as one step when the pointer is released.
 */
export function imageNodeView(cache: ImageMediaCache, options: ImageNodeViewOptions = {}) {
	return (node: ProseMirrorNode, view: EditorView, getPos: () => number | undefined): NodeView => {
		if (node.attrs.shape) return shapeNodeView(node, view, getPos, () => options.theme?.());
		const diagram = parseDiagram(node.attrs.diagram);
		if (diagram) return smartArtNodeView(node, diagram, options.theme?.());
		if (node.attrs.chart) return chartNodeView(node, view, getPos, options);
		const content = renderSpec(node);
		// A placeholder (unsupported drawing or text box) shows static child text that ProseMirror must leave alone.
		if (content.tagName !== 'IMG') return { dom: content, ignoreMutation: () => true };
		const refresh = () => {
			const src =
				(node.attrs.svgPartName && cache.urlFor(String(node.attrs.svgPartName), 'image/svg+xml')) ||
				cache.urlFor(String(node.attrs.partName), String(node.attrs.contentType));
			if (src) content.setAttribute('src', src);
			content.classList.toggle('dve-image-missing', !src);
		};
		refresh();
		const unsubscribe = cache.onChange(refresh);
		const dom = document.createElement('span');
		dom.className = ['dve-picture', placementClass(node.attrs.placement)].filter(Boolean).join(' ');
		dom.append(content);
		content.addEventListener('dblclick', () => {
			const pos = getPos();
			if (pos !== undefined && view.editable) options.editPicture?.(pos);
		});
		const handle = createResizeHandle('dve-picture-handle', {
			...sizeOf(node, view, getPos, options),
			keepRatio: () => true,
			zoomOf: () => dom,
			preview: (width, height) => {
				content.setAttribute('width', String(width));
				content.setAttribute('height', String(height));
			},
		});
		dom.append(handle);
		return {
			dom,
			destroy: unsubscribe,
			selectNode: () => dom.classList.add('dve-picture-selected'),
			deselectNode: () => dom.classList.remove('dve-picture-selected'),
			stopEvent: (event) => event.target === handle,
			ignoreMutation: () => true,
		};
	};
}
