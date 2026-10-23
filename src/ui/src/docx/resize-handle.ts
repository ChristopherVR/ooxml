// The corner resize handle of an inline drawing on the editing surface (pictures, shapes and
// charts): a drag previews the new size and the release commits it once, so a resize is one step
// of the editor's undo history. Pointer moves are divided by the drawing's on-screen zoom.

export interface ResizeHandleOptions {
	/** Whether the document takes edits now. */
	editable(): boolean;
	/** The drawing's size in CSS pixels when the drag starts. */
	size(): { width: number; height: number };
	/** The smallest width and height, in CSS pixels. */
	min: number;
	/** The widest the drawing may become. */
	maxWidth?(): number;
	/** Whether the aspect ratio is kept for this move (pictures always; charts with Shift). */
	keepRatio(event: PointerEvent): boolean;
	/** The element whose drawn and layout widths give the zoom factor. */
	zoomOf(): HTMLElement;
	/** Shows the size being dragged to. */
	preview(width: number, height: number): void;
	/** Applies the final size; only called when it changed. */
	commit(width: number, height: number): void;
}

/** A handle element (`className`) that resizes a drawing by dragging. */
export function createResizeHandle(className: string, options: ResizeHandleOptions): HTMLElement {
	const handle = document.createElement('span');
	handle.className = className;
	handle.setAttribute('aria-hidden', 'true');
	handle.addEventListener('pointerdown', (event) => {
		if (!options.editable()) return;
		event.preventDefault();
		event.stopPropagation();
		const startX = event.clientX;
		const startY = event.clientY;
		const start = options.size();
		const ratio = start.height / Math.max(1, start.width);
		const element = options.zoomOf();
		const zoom = element.getBoundingClientRect().width / Math.max(1, element.offsetWidth) || 1;
		const maxWidth = options.maxWidth?.() ?? Number.POSITIVE_INFINITY;
		let width = start.width;
		let height = start.height;
		const move = (moveEvent: PointerEvent) => {
			width = Math.min(
				maxWidth,
				Math.max(options.min, start.width + (moveEvent.clientX - startX) / zoom),
			);
			height = options.keepRatio(moveEvent)
				? Math.max(options.min, width * ratio)
				: Math.max(options.min, start.height + (moveEvent.clientY - startY) / zoom);
			options.preview(Math.round(width), Math.round(height));
		};
		const up = () => {
			window.removeEventListener('pointermove', move);
			window.removeEventListener('pointerup', up);
			const w = Math.round(width);
			const h = Math.round(height);
			if (w !== start.width || h !== start.height) options.commit(w, h);
		};
		window.addEventListener('pointermove', move);
		window.addEventListener('pointerup', up);
	});
	return handle;
}
