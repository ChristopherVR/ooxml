// How a drawing shape looks on the editing surface: its geometry as SVG at its size, placed where its
// wrapping puts it, with a resize handle while selected.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';
import type { DocumentModel, ShapeSpec, ThemeCatalog } from 'ooxml-core/docx';
import { createResizeHandle } from './resize-handle';
import { paintShape } from './story-render';
import type { PictureMedia } from './picture-paint';

const MIN_SIZE = 4;

function parseSpec(value: unknown): ShapeSpec | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as ShapeSpec;
	} catch {
		return undefined;
	}
}

export function shapeNodeView(
	initial: ProseMirrorNode,
	view: EditorView,
	getPos: () => number | undefined,
	theme: () => ThemeCatalog | undefined,
	model: () => DocumentModel | undefined = () => undefined,
	media?: PictureMedia,
): NodeView {
	let node = initial;
	let selected = false;
	const dom = document.createElement('span');
	const holder = document.createElement('span');
	holder.className = 'dve-shape-drawing';
	const handle = createResizeHandle('dve-shape-handle', {
		// Only a size the file can take: a shape the model writes, or a simple text box.
		editable: () =>
			view.editable && Boolean(node.attrs.textBoxEditable || parseSpec(node.attrs.shape)?.editable),
		size: () => ({ width: Number(node.attrs.widthPx), height: Number(node.attrs.heightPx) }),
		min: MIN_SIZE,
		keepRatio: () => false,
		zoomOf: () => dom,
		preview: (width, height) => {
			dom.style.width = `${width}px`;
			dom.style.height = `${height}px`;
		},
		commit: (width, height) => {
			const pos = getPos();
			if (pos === undefined) return;
			view.dispatch(
				view.state.tr.setNodeMarkup(pos, undefined, {
					...node.attrs,
					widthPx: width,
					heightPx: height,
				}),
			);
		},
	});
	dom.append(holder, handle);
	dom.contentEditable = 'false';

	const pictureUrl = media && ((part: string, type: string) => media.urlFor(part, type));
	const render = () => paintShape(dom, holder, node, theme(), model(), selected, pictureUrl);
	render();
	// Pictures inside a text box repaint when late bytes replace the cache's URLs.
	const unsubscribe = media?.onChange(render);

	return {
		dom,
		update(next) {
			if (next.type !== node.type || !next.attrs.shape) return false;
			node = next;
			render();
			return true;
		},
		selectNode: () => {
			selected = true;
			dom.classList.add('dve-shape-selected');
		},
		deselectNode: () => {
			selected = false;
			dom.classList.remove('dve-shape-selected');
		},
		stopEvent: (event) => event.target === handle,
		ignoreMutation: () => true,
		...(unsubscribe ? { destroy: unsubscribe } : {}),
	};
}
