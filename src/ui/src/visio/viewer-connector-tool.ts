import type { VisioPage } from 'ooxml-core/visio';
import {
	editErrorMessage,
	isEditCancellation,
	visioConnectorCreationCommand,
	visioDrawIsLargeEnough,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { ViewerDrawingGesture } from './viewer-drawing-gesture';

const TARGET = 'connectTarget';

/** A page shape the core can glue to: a visible, local, top-level 2D shape. */
export function connectorGlueTarget(page: VisioPage, element: Element | null): string | undefined {
	let group = element?.closest?.<SVGGElement>('[data-shape-id]');
	while (group) {
		const id = group.dataset.shapeId,
			shape = page.shapes.find((candidate) => candidate.id === id);
		if (shape)
			return shape.kind === 'shape' && !shape.hidden && !shape.masterId && !shape.children.length
				? shape.id
				: undefined;
		group = group.parentElement?.closest<SVGGElement>('[data-shape-id]');
	}
	return undefined;
}

/**
 * Home > Tools > Connector (Ctrl+3). A drag draws a straight dynamic connector; ends released on
 * shapes are glued to them, so the connector follows when they move. Shapes under the pointer are
 * highlighted as glue targets. Routing is straight; right-angle routing is not performed.
 */
export class ViewerConnectorTool {
	#gesture: ViewerDrawingGesture;
	#request = 0;
	#begin: string | undefined;
	#hover: string | undefined;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: { active(): boolean; announce(message: string): void },
	) {
		this.#gesture = new ViewerDrawingGesture(viewport, controller, {
			tool: () => (options.active() ? 'line' : undefined),
			announce: options.announce,
			finish: async (drag, end) => {
				const request = ++this.#request;
				const glue = { begin: this.#begin, end: this.#hover };
				try {
					if (!visioDrawIsLargeEnough('line', drag.start, end) || !this.#gesture.current(drag))
						return;
					const command = visioConnectorCreationCommand(drag.page, drag.start, end, {
						...(glue.begin === undefined ? {} : { begin: glue.begin }),
						...(glue.end === undefined ? {} : { end: glue.end }),
					});
					await controller.applyCreationEdits([command], drag.token);
					if (request === this.#request)
						options.announce(
							`Connector ${command.shapeId} added${command.connect?.begin || command.connect?.end ? ' and glued' : ''}.`,
						);
				} catch (error) {
					if (request === this.#request && !isEditCancellation(error))
						options.announce(editErrorMessage(error));
				} finally {
					this.#highlight(undefined);
				}
			},
		});
	}
	get drawing(): boolean {
		return this.#gesture.drawing;
	}
	/** The shape currently highlighted as a glue target, for tests and status text. */
	get target(): string | undefined {
		return this.#hover;
	}
	cancel(): void {
		++this.#request;
		this.#gesture.cancel();
		this.#begin = undefined;
		this.#highlight(undefined);
	}
	render(state: ViewerState): void {
		this.#gesture.render(state);
		if (!this.options.active()) this.#highlight(undefined);
		else if (this.#hover) this.#highlight(this.#hover);
	}
	wire(): () => void {
		const Abort = this.viewport.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal, capture: true };
		// Registered before the gesture so the begin target is known when the drag starts.
		this.viewport.addEventListener(
			'pointerdown',
			(event) => {
				if (!this.options.active() || this.#gesture.drawing) return;
				this.#begin = this.#targetAt(event);
			},
			options,
		);
		this.viewport.addEventListener(
			'pointermove',
			(event) => {
				if (this.options.active()) this.#highlight(this.#targetAt(event));
			},
			options,
		);
		this.viewport.addEventListener(
			'pointerleave',
			() => {
				if (!this.#gesture.drawing) this.#highlight(undefined);
			},
			options,
		);
		const dispose = this.#gesture.wire();
		return () => {
			this.cancel();
			dispose();
			events.abort();
		};
	}
	#page(): VisioPage | undefined {
		const state = this.controller.state;
		return state.document?.pages[state.pageIndex];
	}
	/** Pointer capture retargets events to the viewport, so hit-test the point itself. */
	#targetAt(event: PointerEvent): string | undefined {
		const page = this.#page();
		if (!page) return undefined;
		const root = this.viewport.getRootNode() as Document | ShadowRoot;
		const hits =
			typeof root.elementsFromPoint === 'function'
				? root.elementsFromPoint(event.clientX, event.clientY)
				: [event.target as Element];
		for (const hit of hits) {
			if (!this.viewport.contains(hit) || hit.classList?.contains('draw-preview')) continue;
			const id = connectorGlueTarget(page, hit);
			if (id) return id;
		}
		return undefined;
	}
	#highlight(id: string | undefined): void {
		if (id === undefined && this.#hover === undefined) return;
		this.#hover = id;
		for (const group of this.viewport.querySelectorAll<SVGGElement>('[data-connect-target]'))
			if (group.dataset.shapeId !== id) delete group.dataset[TARGET];
		if (id)
			for (const group of this.viewport.querySelectorAll<SVGGElement>('[data-shape-id]'))
				if (group.dataset.shapeId === id) group.dataset[TARGET] = 'true';
	}
}
