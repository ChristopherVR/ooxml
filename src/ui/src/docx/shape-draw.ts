// Drawing a shape by dragging on the page: the armed preset, the rubber band and the insertion.
import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { isDirectionalPreset } from 'ooxml-core/docx';
import { insertDrawnShape } from './shape-commands';

/** Which shape the next drag on the page draws (set from Insert > Shapes), shared by the editor. */
export class ShapeTool {
	private current: string | undefined;
	private readonly listeners = new Set<() => void>();

	get preset(): string | undefined {
		return this.current;
	}
	arm(preset: string): void {
		this.current = preset;
		for (const listener of this.listeners) listener();
	}
	disarm(): void {
		if (this.current === undefined) return;
		this.current = undefined;
		for (const listener of this.listeners) listener();
	}
	onChange(listener: () => void): () => void {
		this.listeners.add(listener);
		return () => {
			this.listeners.delete(listener);
		};
	}
}

const DRAG_THRESHOLD = 5;

/** Ratio between on-screen and layout pixels (the editor's zoom). */
const scaleOf = (element: HTMLElement): number =>
	element.offsetWidth ? element.getBoundingClientRect().width / element.offsetWidth || 1 : 1;

function paragraphAt(
	view: EditorView,
	pos: number,
): { element: HTMLElement; start: number } | undefined {
	const $pos = view.state.doc.resolve(Math.min(pos, view.state.doc.content.size));
	for (let depth = $pos.depth; depth > 0; depth--) {
		if ($pos.node(depth).type.name !== 'paragraph') continue;
		const dom = view.nodeDOM($pos.before(depth));
		return dom instanceof HTMLElement ? { element: dom, start: $pos.start(depth) } : undefined;
	}
	return undefined;
}

function beginDrag(view: EditorView, tool: ShapeTool, preset: string, event: MouseEvent): void {
	const hit = view.posAtCoords({ left: event.clientX, top: event.clientY });
	const paragraph = hit && paragraphAt(view, hit.pos);
	if (!hit || !paragraph) {
		tool.disarm();
		return;
	}
	const startX = event.clientX;
	const startY = event.clientY;
	const band = document.createElement('div');
	band.className = 'dve-shape-rubber';
	band.hidden = true;
	const root = view.dom.getRootNode();
	(root instanceof ShadowRoot ? root : document.body).append(band);
	const rect = (clientX: number, clientY: number) => ({
		left: Math.min(startX, clientX),
		top: Math.min(startY, clientY),
		width: Math.abs(clientX - startX),
		height: Math.abs(clientY - startY),
	});
	const move = (moveEvent: MouseEvent) => {
		const box = rect(moveEvent.clientX, moveEvent.clientY);
		band.hidden = box.width < DRAG_THRESHOLD && box.height < DRAG_THRESHOLD;
		Object.assign(band.style, {
			left: `${box.left}px`,
			top: `${box.top}px`,
			width: `${box.width}px`,
			height: `${box.height}px`,
		});
	};
	const finish = (upEvent: MouseEvent | undefined) => {
		window.removeEventListener('mousemove', move, true);
		window.removeEventListener('mouseup', finish, true);
		window.removeEventListener('keydown', cancel, true);
		band.remove();
		if (!upEvent) return;
		const box = rect(upEvent.clientX, upEvent.clientY);
		const bounds = paragraph.element.getBoundingClientRect();
		const scale = scaleOf(paragraph.element);
		const dragged = box.width >= DRAG_THRESHOLD || box.height >= DRAG_THRESHOLD;
		const directional = isDirectionalPreset(preset);
		insertDrawnShape(view, preset, paragraph.start, {
			x: (box.left - bounds.left) / scale,
			y: (box.top - bounds.top) / scale,
			...(dragged
				? {
						widthPx: box.width / scale,
						heightPx: directional ? box.height / scale : Math.max(box.height / scale, 4),
					}
				: {}),
			...(dragged && directional && upEvent.clientX < startX ? { flipH: true } : {}),
			...(dragged && directional && upEvent.clientY < startY ? { flipV: true } : {}),
		});
		tool.disarm();
		view.focus();
	};
	const cancel = (keyEvent: KeyboardEvent) => {
		if (keyEvent.key !== 'Escape') return;
		finish(undefined);
		tool.disarm();
	};
	window.addEventListener('mousemove', move, true);
	window.addEventListener('mouseup', finish, true);
	window.addEventListener('keydown', cancel, true);
}

/** The plugin that turns a drag on the page into a shape while `tool` is armed. */
export function shapeDrawPlugin(tool: ShapeTool): Plugin {
	return new Plugin({
		view: (view) => {
			const sync = () => view.dom.classList.toggle('dve-drawing-shape', tool.preset !== undefined);
			sync();
			const off = tool.onChange(sync);
			return {
				destroy: () => {
					off();
					view.dom.classList.remove('dve-drawing-shape');
				},
			};
		},
		props: {
			handleDOMEvents: {
				mousedown: (view, event) => {
					const preset = tool.preset;
					if (!preset || !view.editable || event.button !== 0) return false;
					event.preventDefault();
					beginDrag(view, tool, preset, event);
					return true;
				},
			},
			handleKeyDown: (_view, event) => {
				if (event.key !== 'Escape' || !tool.preset) return false;
				tool.disarm();
				return true;
			},
		},
	});
}
