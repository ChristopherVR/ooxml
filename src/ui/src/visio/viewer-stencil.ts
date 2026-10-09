import type { ViewerController } from './controller';
import { editErrorMessage, isEditCancellation } from 'ooxml-core/visio/ui';
import { MASTER_MIME, masterCreation } from './shapes-window';
import { insertMaster, pagePoint } from './viewer-draw-tool';

/**
 * Shapes window interaction: drag a master onto the page to drop it there, or activate it to add
 * it at the page centre. Ellipse and Circle are native ellipses; the rest are core outlines.
 */
export function wireStencil(
	pane: HTMLElement,
	viewport: HTMLElement,
	controller: ViewerController,
	announce: (message: string) => void,
): () => void {
	const Abort = pane.ownerDocument.defaultView?.AbortController ?? AbortController;
	const events = new Abort();
	const options = { signal: events.signal };
	// A drop is accepted while an earlier edit is still saving; the shape is added once it is done.
	const editable = () => {
		const { edit, loading } = controller.state;
		return edit.sourceAvailable && !loading;
	};
	const idle = () =>
		new Promise<void>((resolve) => {
			let stop: (() => void) | undefined;
			let done = false;
			stop = controller.subscribe((state) => {
				if (done || state.edit.busy) return;
				done = true;
				resolve();
				stop?.();
			});
			if (done) stop();
		});
	// Drops are added one at a time, in the order they were made, each where it was released.
	let queue = Promise.resolve();
	const add = (id: string, centre?: { clientX: number; clientY: number }) => {
		const master = masterCreation(id);
		const state = controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!master || !page) return queue;
		if (!editable()) {
			announce('Open a .vsdx file to add shapes. Model-only documents are read only.');
			return queue;
		}
		const { size } = master;
		const svg = viewport.querySelector<SVGSVGElement>('svg.paper');
		const point = centre && svg ? pagePoint(svg, page, centre) : undefined;
		// Keep the whole shape on the page, as the drop point is its centre.
		const x = Math.min(
			page.width - size.width / 2,
			Math.max(size.width / 2, point?.x ?? page.width / 2),
		);
		const y = Math.min(
			page.height - size.height / 2,
			Math.max(size.height / 2, point?.y ?? page.height / 2),
		);
		queue = queue.then(async () => {
			try {
				if (controller.state.edit.busy) await idle();
				const current = controller.state;
				const target = current.document?.pages[current.pageIndex];
				if (!target || target.id !== page.id || !editable()) return;
				const shapeId = await insertMaster(controller, target, master.create, { x, y }, size);
				const name = pane.querySelector(`[data-master="${id}"] span`)?.textContent ?? 'Shape';
				announce(`${name} ${shapeId} added from Basic Shapes.`);
			} catch (error) {
				if (!isEditCancellation(error) && !controller.state.edit.error)
					announce(editErrorMessage(error));
			}
		});
		return queue;
	};
	pane.addEventListener(
		'dragstart',
		(event) => {
			const master = (event.target as Element).closest?.<HTMLElement>('[data-master]');
			if (!master || !masterCreation(master.dataset.master!) || !event.dataTransfer) return;
			event.dataTransfer.setData(MASTER_MIME, master.dataset.master!);
			event.dataTransfer.effectAllowed = 'copy';
		},
		options,
	);
	pane.addEventListener(
		'click',
		(event) => {
			const master = (event.target as Element).closest?.<HTMLElement>('[data-master]');
			if (master) void add(master.dataset.master!);
		},
		options,
	);
	viewport.addEventListener(
		'dragover',
		(event) => {
			if (!event.dataTransfer?.types.includes(MASTER_MIME) || !editable()) return;
			event.preventDefault();
			event.dataTransfer.dropEffect = 'copy';
		},
		options,
	);
	viewport.addEventListener(
		'drop',
		(event) => {
			const id = event.dataTransfer?.getData(MASTER_MIME);
			if (!id) return;
			event.preventDefault();
			void add(id, event);
		},
		options,
	);
	return () => events.abort();
}
