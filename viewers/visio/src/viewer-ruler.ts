/**
 * Visio's rulers: inch scales along the top and left of the drawing window with 0 at the page's
 * left and bottom edges (vertical values grow upward, as in Visio). Viewer-only presentation:
 * nothing is read from or written to the document beyond the rendered page box.
 */
const SIZE = 18;

export interface Rulers {
	element: HTMLElement;
	render(visible: boolean, zoom: number): void;
	wire(): () => void;
}

/** Wraps the viewport in a grid with a corner, a top ruler and a left ruler. */
export function createRulers(viewport: HTMLElement): Rulers {
	const doc = viewport.ownerDocument;
	const area = doc.createElement('div');
	area.className = 'canvas-area';
	const corner = doc.createElement('div');
	corner.className = 'ruler-corner';
	const top = doc.createElement('canvas');
	top.className = 'ruler ruler-h';
	const left = doc.createElement('canvas');
	left.className = 'ruler ruler-v';
	for (const ruler of [corner, top, left]) ruler.setAttribute('aria-hidden', 'true');
	viewport.replaceWith(area);
	area.append(corner, top, left, viewport);
	let visible = false;
	let zoom = 1;
	let frame = 0;
	const draw = () => {
		frame = 0;
		if (!visible) return;
		const paper = viewport.querySelector('svg.paper');
		const ratio = doc.defaultView?.devicePixelRatio ?? 1;
		const style = doc.defaultView?.getComputedStyle(area);
		const ink = style?.getPropertyValue('--_vv-muted').trim() || '#605e5c';
		const base = viewport.getBoundingClientRect();
		const page = paper?.getBoundingClientRect();
		const inch = 96 * zoom;
		// Finer ticks as the zoom grows, like Visio: halves, quarters, eighths, sixteenths.
		const divisions = inch >= 192 ? 16 : inch >= 96 ? 8 : inch >= 48 ? 4 : 2;
		const paint = (canvas: HTMLCanvasElement, horizontal: boolean) => {
			const length = horizontal ? base.width : base.height;
			canvas.width = Math.max(1, Math.round(length * ratio));
			canvas.height = Math.round(SIZE * ratio);
			if (!horizontal) [canvas.width, canvas.height] = [canvas.height, canvas.width];
			const context = canvas.getContext?.('2d');
			if (!context || !page) return;
			context.scale(ratio, ratio);
			context.strokeStyle = ink;
			context.fillStyle = ink;
			context.lineWidth = 1;
			context.font = '9px "Segoe UI", Arial, sans-serif';
			// Pixel offset of the page origin (left edge, or bottom edge for the vertical ruler).
			const origin = horizontal ? page.left - base.left : page.bottom - base.top;
			const step = inch / divisions;
			const first = Math.floor((horizontal ? -origin : origin - length) / step) - 1;
			const last = Math.ceil((horizontal ? length - origin : origin) / step) + 1;
			context.beginPath();
			for (let index = first; index <= last; index++) {
				const position = horizontal ? origin + index * step : origin - index * step;
				const tick =
					index % divisions === 0 ? SIZE : index % (divisions / 2) === 0 ? SIZE * 0.5 : SIZE * 0.3;
				const at = Math.round(position) + 0.5;
				if (horizontal) {
					context.moveTo(at, SIZE);
					context.lineTo(at, SIZE - tick);
				} else {
					context.moveTo(SIZE, at);
					context.lineTo(SIZE - tick, at);
				}
				if (index % divisions === 0) {
					const label = String(index / divisions);
					if (horizontal) context.fillText(label, at + 2, 9);
					else {
						context.save();
						context.translate(9, at - 2);
						context.rotate(-Math.PI / 2);
						context.fillText(label, 0, 0);
						context.restore();
					}
				}
			}
			context.stroke();
		};
		paint(top, true);
		paint(left, false);
	};
	const schedule = () => {
		if (!visible || frame) return;
		const view = doc.defaultView;
		frame = view?.requestAnimationFrame?.(draw) ?? 0;
		if (!frame) draw();
	};
	return {
		element: area,
		render(nextVisible, nextZoom) {
			visible = nextVisible;
			zoom = nextZoom;
			area.dataset.ruler = String(visible);
			schedule();
		},
		wire() {
			const Abort = doc.defaultView?.AbortController ?? AbortController;
			const events = new Abort();
			viewport.addEventListener('scroll', schedule, { signal: events.signal, passive: true });
			const Observer = doc.defaultView?.ResizeObserver;
			const observer = Observer ? new Observer(schedule) : undefined;
			observer?.observe(viewport);
			return () => {
				events.abort();
				observer?.disconnect();
				if (frame) doc.defaultView?.cancelAnimationFrame?.(frame);
				frame = 0;
			};
		},
	};
}
