import type { ViewerController } from './controller.js';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** Visio's Pan & Zoom window: a floating page thumbnail with the visible area outlined. */
export function createPanZoom(doc: Document): HTMLElement {
	const pane = doc.createElement('section');
	pane.className = 'pan-zoom';
	pane.hidden = true;
	pane.setAttribute('aria-label', 'Pan & Zoom');
	const heading = doc.createElement('div');
	heading.className = 'pan-zoom-heading';
	const title = doc.createElement('span');
	title.textContent = 'Pan & Zoom';
	const close = doc.createElement('button');
	close.type = 'button';
	close.className = 'pan-zoom-close';
	close.dataset.chrome = 'pan-zoom';
	close.setAttribute('aria-label', 'Close Pan & Zoom');
	close.textContent = '×';
	heading.append(title, close);
	const view = doc.createElement('div');
	view.className = 'pan-zoom-view';
	view.tabIndex = 0;
	view.setAttribute('role', 'application');
	view.setAttribute(
		'aria-label',
		'Page overview. Click or drag to pan; arrow keys move the view; Ctrl and the mouse wheel zoom.',
	);
	pane.append(heading, view);
	return pane;
}

/**
 * Keeps the thumbnail and the visible-area rectangle in step with the drawing window, and pans
 * the drawing window from the thumbnail. The thumbnail is a DOM clone of the rendered page.
 */
export class ViewerPanZoom {
	readonly #pane: HTMLElement;
	readonly #view: HTMLElement;
	#paper: Element | null = null;
	#frame: SVGRectElement | null = null;
	#raf = 0;
	constructor(
		root: ShadowRoot,
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly checked: (open: boolean) => void,
	) {
		this.#pane = root.querySelector('.pan-zoom')!;
		this.#view = this.#pane.querySelector('.pan-zoom-view')!;
	}
	get open(): boolean {
		return !this.#pane.hidden;
	}
	toggle(open = !this.open): void {
		this.#pane.hidden = !open;
		this.checked(open);
		if (open) {
			this.#paper = null;
			this.render();
		}
	}
	wire(): () => void {
		const doc = this.viewport.ownerDocument;
		const Abort = doc.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.#pane.addEventListener(
			'click',
			(event) => {
				if ((event.target as Element).closest?.('.pan-zoom-close')) this.toggle(false);
			},
			options,
		);
		this.viewport.addEventListener('scroll', () => this.#schedule(), { ...options, passive: true });
		let dragging = false;
		const panTo = (event: PointerEvent) => {
			const thumb = this.#view.querySelector('svg');
			const page = this.#page();
			if (!thumb || !page) return;
			const box = thumb.getBoundingClientRect();
			// Centre the drawing window on the pointed page position.
			this.#centre(
				((event.clientX - box.left) / box.width) * page.width,
				((event.clientY - box.top) / box.height) * page.height,
			);
		};
		this.#view.addEventListener(
			'pointerdown',
			(event) => {
				if (event.button !== 0) return;
				dragging = true;
				this.#view.setPointerCapture?.(event.pointerId);
				panTo(event);
			},
			options,
		);
		this.#view.addEventListener('pointermove', (event) => dragging && panTo(event), options);
		for (const type of ['pointerup', 'pointercancel'] as const)
			this.#view.addEventListener(type, () => (dragging = false), options);
		this.#view.addEventListener(
			'wheel',
			(event) => {
				if (!event.ctrlKey && !event.metaKey) return;
				event.preventDefault();
				const zoom = this.controller.state.zoom;
				this.controller.setZoom(event.deltaY < 0 ? zoom * 1.1 : zoom / 1.1);
			},
			{ ...options, passive: false },
		);
		this.#view.addEventListener(
			'keydown',
			(event) => {
				const steps: Record<string, [number, number]> = {
					ArrowLeft: [-1, 0],
					ArrowRight: [1, 0],
					ArrowUp: [0, -1],
					ArrowDown: [0, 1],
				};
				const step = steps[event.key];
				if (!step) return;
				event.preventDefault();
				this.viewport.scrollBy({ left: step[0] * 48, top: step[1] * 48 });
			},
			options,
		);
		return () => {
			events.abort();
			if (this.#raf) doc.defaultView?.cancelAnimationFrame?.(this.#raf);
		};
	}
	#page() {
		const { document, pageIndex } = this.controller.state;
		return document?.pages[pageIndex];
	}
	#centre(x: number, y: number): void {
		const paper = this.viewport.querySelector('svg.paper');
		const page = this.#page();
		if (!paper || !page) return;
		const box = paper.getBoundingClientRect();
		const port = this.viewport.getBoundingClientRect();
		const scale = box.width / page.width;
		this.viewport.scrollBy({
			left: box.left + x * scale - (port.left + port.width / 2),
			top: box.top + y * scale - (port.top + port.height / 2),
		});
	}
	#schedule(): void {
		if (!this.open || this.#raf) return;
		const view = this.viewport.ownerDocument.defaultView;
		this.#raf =
			view?.requestAnimationFrame?.(() => {
				this.#raf = 0;
				this.render();
			}) ?? 0;
		if (!this.#raf) this.render();
	}
	/** Refresh the thumbnail when the page changed, then the visible-area rectangle. */
	render(): void {
		if (!this.open) return;
		const paper = this.viewport.querySelector('svg.paper');
		const page = this.#page();
		if (!paper || !page) {
			this.#view.replaceChildren();
			this.#paper = null;
			return;
		}
		if (paper !== this.#paper) {
			this.#paper = paper;
			const thumb = paper.cloneNode(true) as SVGSVGElement;
			thumb.removeAttribute('style');
			thumb.removeAttribute('class');
			thumb.setAttribute('aria-hidden', 'true');
			thumb.querySelectorAll('[tabindex]').forEach((node) => node.removeAttribute('tabindex'));
			this.#frame = this.viewport.ownerDocument.createElementNS(SVG_NS, 'rect');
			this.#frame.classList.add('pan-zoom-frame');
			this.#frame.setAttribute('vector-effect', 'non-scaling-stroke');
			thumb.append(this.#frame);
			this.#view.replaceChildren(thumb);
		}
		const box = paper.getBoundingClientRect();
		const port = this.viewport.getBoundingClientRect();
		if (!box.width || !this.#frame) return;
		const scale = page.width / box.width;
		const left = Math.max(0, (port.left - box.left) * scale);
		const top = Math.max(0, (port.top - box.top) * scale);
		const right = Math.min(page.width, (port.right - box.left) * scale);
		const bottom = Math.min(page.height, (port.bottom - box.top) * scale);
		this.#frame.setAttribute('x', String(left));
		this.#frame.setAttribute('y', String(top));
		this.#frame.setAttribute('width', String(Math.max(0, right - left)));
		this.#frame.setAttribute('height', String(Math.max(0, bottom - top)));
	}
}
