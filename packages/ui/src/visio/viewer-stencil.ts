import type { ViewerController } from './controller.js';
import { editErrorMessage, isEditCancellation } from './edit-error.js';
import { MASTER_MIME, masterSize } from './shapes-window.js';
import { insertRectangle, pagePoint } from './viewer-draw-tool.js';

/**
 * Shapes window interaction: drag a master onto the page to drop it there, or activate it to add
 * it at the page centre. Only masters the core can create are draggable; the rest stay inert.
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
	const editable = () => {
		const { edit, loading } = controller.state;
		return edit.sourceAvailable && !edit.busy && !loading;
	};
	const add = async (id: string, centre?: { clientX: number; clientY: number }) => {
		const size = masterSize(id);
		const state = controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!size || !page) return;
		if (!editable()) {
			announce('Open a .vsdx file to add shapes. Model-only documents are read only.');
			return;
		}
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
		try {
			const shapeId = await insertRectangle(controller, page, { x, y }, size);
			announce(`Rectangle ${shapeId} added from Basic Shapes.`);
		} catch (error) {
			if (!isEditCancellation(error) && !controller.state.edit.error)
				announce(editErrorMessage(error));
		}
	};
	pane.addEventListener(
		'dragstart',
		(event) => {
			const master = (event.target as Element).closest?.<HTMLElement>('[data-master]');
			if (!master || !masterSize(master.dataset.master!) || !event.dataTransfer) return;
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
