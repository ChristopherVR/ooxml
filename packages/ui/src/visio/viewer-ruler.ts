/**
 * Visio's rulers: inch scales along the top and left of the drawing window with 0 at the page's
 * left and bottom edges (vertical values grow upward, as in Visio). Drawn by the shared
 * `office-ui-ruler`; this module only lays them out and says where zero is. Viewer-only
 * presentation: nothing is read from or written to the document beyond the rendered page box.
 */
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
	corner.setAttribute('aria-hidden', 'true');
	const top = doc.createElement('office-ui-ruler');
	top.className = 'ruler ruler-h';
	const left = doc.createElement('office-ui-ruler');
	left.className = 'ruler ruler-v';
	left.setAttribute('orientation', 'vertical');
	left.setAttribute('direction', 'reverse');
	viewport.replaceWith(area);
	area.append(corner, top, left, viewport);
	let visible = false;
	let zoom = 1;
	let frame = 0;
	const place = () => {
		frame = 0;
		if (!visible) return;
		const paper = viewport.querySelector('svg.paper');
		if (!paper) return;
		const base = viewport.getBoundingClientRect();
		const page = paper.getBoundingClientRect();
		const scale = String(96 * zoom);
		top.setAttribute('scale', scale);
		left.setAttribute('scale', scale);
		top.setAttribute('origin', String(page.left - base.left));
		left.setAttribute('origin', String(page.bottom - base.top));
	};
	const schedule = () => {
		if (!visible || frame) return;
		frame = doc.defaultView?.requestAnimationFrame?.(place) ?? 0;
		if (!frame) place();
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
