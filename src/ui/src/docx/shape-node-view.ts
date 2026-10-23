// How a drawing shape looks on the editing surface: its geometry as SVG at its size, placed where its
// wrapping puts it, with a resize handle while selected.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';
import {
	shapeSvg,
	type PicturePlacement,
	type ShapeSpec,
	type ThemeCatalog,
} from 'ooxml-core/docx';
import { placementClass } from 'ooxml-core/docx/ui';
import { createResizeHandle } from './resize-handle';

const MIN_SIZE = 4;

function parse<T>(value: unknown): T | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as T;
	} catch {
		return undefined;
	}
}

/** Whether a floating shape in front of or behind text can sit at its offsets inside its paragraph. */
function absolute(placement: PicturePlacement | undefined): placement is PicturePlacement {
	return (
		placement?.wrap === 'none' &&
		(placement.relativeFromV === undefined ||
			placement.relativeFromV === 'paragraph' ||
			placement.relativeFromV === 'line') &&
		placement.align === undefined &&
		placement.alignV === undefined
	);
}

export function shapeNodeView(
	initial: ProseMirrorNode,
	view: EditorView,
	getPos: () => number | undefined,
	theme: () => ThemeCatalog | undefined,
): NodeView {
	let node = initial;
	let selected = false;
	const dom = document.createElement('span');
	const holder = document.createElement('span');
	holder.className = 'dve-shape-drawing';
	const handle = createResizeHandle('dve-shape-handle', {
		editable: () => view.editable,
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

	const render = () => {
		const spec = parse<ShapeSpec>(node.attrs.shape);
		const placement = parse<PicturePlacement>(node.attrs.placement);
		const width = Number(node.attrs.widthPx);
		const height = Number(node.attrs.heightPx);
		const anchored = Boolean(node.attrs.anchored);
		const classes = ['dve-shape', selected ? 'dve-shape-selected' : ''];
		if (anchored && absolute(placement)) {
			classes.push('dve-shape-floating');
			if (placement.behindText) classes.push('dve-shape-behind');
			dom.style.left = `${placement.offsetXPx ?? 0}px`;
			dom.style.top = `${placement.offsetYPx ?? 0}px`;
		} else {
			dom.style.left = '';
			dom.style.top = '';
			if (anchored) classes.push(placementClass(JSON.stringify(placement ?? {})));
		}
		dom.className = classes.filter(Boolean).join(' ');
		dom.dataset.shapePreset = spec?.preset ?? '';
		dom.style.width = `${width}px`;
		dom.style.height = `${height}px`;
		const svg = spec && shapeSvg(spec, width, height, theme());
		holder.replaceChildren();
		if (svg) {
			const parsed = new DOMParser().parseFromString(svg, 'image/svg+xml').documentElement;
			if (parsed.localName === 'svg') holder.append(document.importNode(parsed, true));
		} else holder.textContent = String(node.attrs.unsupported ?? 'Shape');
	};
	render();

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
	};
}
