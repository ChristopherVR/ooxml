// Collaborators' selections in the grid: the core lays them out (`layoutRemoteSelections`), this
// places one coloured outline with a name tag per peer in every quadrant plane, after each paint.
import { layoutRemoteSelections, type RemoteSelectionBox } from 'ooxml-core/xlsx';
import { QUADRANTS } from 'ooxml-core/xlsx/ui';
import { h, place } from './dom';
import type { GridView } from './grid-view';

const TAG_HEIGHT = 16;

class RemoteLayer {
	readonly element: HTMLDivElement;
	readonly #nodes: HTMLDivElement[] = [];

	constructor(private readonly doc: Document) {
		this.element = h(doc, 'div', 'xg-remote-layer');
	}

	render(boxes: readonly RemoteSelectionBox[]): void {
		boxes.forEach((box, index) => {
			let node = this.#nodes[index];
			if (!node) {
				node = h(this.doc, 'div', 'xg-remote');
				node.append(h(this.doc, 'span', 'xg-remote-tag'));
				this.#nodes.push(node);
				this.element.append(node);
			}
			node.hidden = false;
			place(node, box.x - 1, box.y - 1, box.w + 1, box.h + 1);
			node.style.borderColor = box.color;
			node.dataset.clientId = String(box.clientId);
			node.dataset.tag = box.tag;
			node.title = box.name;
			const tag = node.firstElementChild as HTMLElement;
			tag.textContent = box.label;
			tag.style.backgroundColor = box.color;
		});
		for (let i = boxes.length; i < this.#nodes.length; i++) {
			const node = this.#nodes[i];
			if (node) node.hidden = true;
		}
	}
}

/** Adds the remote-selection layer to every quadrant; returns the remover. */
export function installRemoteSelections(view: GridView): () => void {
	const ctx = view.ctx;
	const layers = QUADRANTS.map((q) => {
		const layer = new RemoteLayer(view.doc);
		view.quads[q].plane.append(layer.element);
		return layer;
	});
	const hook = (): void => {
		const peers = ctx.remoteSelections?.() ?? [];
		const boxes = peers.length
			? layoutRemoteSelections(view.metrics, peers, {
					sheet: view.sheetIndex(),
					maxRow: view.extentRow,
					maxCol: view.extentCol,
					tagHeight: TAG_HEIGHT,
				})
			: [];
		for (const layer of layers) layer.render(boxes);
	};
	view.hooks.add(hook);
	const stop = ctx.onRemoteSelectionsChange?.(() => view.schedule());
	hook();
	return () => {
		stop?.();
		view.hooks.delete(hook);
		for (const layer of layers) layer.element.remove();
	};
}
