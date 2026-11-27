// The lasso's selection: the strokes it holds in each ink drawing, framed on their pages. It
// belongs to one document: any change (an edit, an undo, a remote change) drops it, and each use
// checks that its drawings are still the ones it framed.
import type { InkCoordinate } from 'ooxml-core/drawingml';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { inkTools } from '../ink/tools';
import { localPoint, pageMapping, stillThere, type InkTarget } from './draw-targets';

export class InkLasso {
	#targets: InkTarget[] = [];
	#boxes: HTMLElement[] = [];
	#doc: ProseMirrorNode | undefined;

	get size(): number {
		return this.#targets.length;
	}

	/** Holds the strokes of `candidates` inside `polygon` (client points) and frames them. */
	select(view: EditorView, candidates: InkTarget[], polygon: InkCoordinate[]): void {
		this.clear();
		this.#targets = candidates.flatMap((target) => {
			const indexes = inkTools().lassoInkStrokes(
				target.ink.strokes,
				polygon.map((point) => localPoint(target, point).point),
			);
			return indexes.length ? [{ ...target, indexes }] : [];
		});
		this.#doc = this.#targets.length ? view.state.doc : undefined;
		for (const target of this.#targets) this.#frame(target);
	}

	/** The held strokes whose drawings are unchanged; the selection is dropped if the document moved on. */
	targets(view: EditorView): InkTarget[] {
		if (this.#doc && view.state.doc !== this.#doc) this.clear();
		return this.#targets.filter((target) => stillThere(view, target));
	}

	/** Drops a selection the document has moved on from (true when it dropped one). */
	follow(doc: ProseMirrorNode | undefined): boolean {
		if (!this.#doc || doc === this.#doc) return false;
		this.clear();
		return true;
	}

	#frame(target: InkTarget): void {
		const bounds = inkTools().inkStrokesBounds(
			target.indexes.map((index) => target.ink.strokes[index]!),
		);
		if (!bounds) return;
		const { ox, oy, kx, ky } = pageMapping(target);
		const frame = document.createElement('div');
		frame.className = 'dve-ink-lasso-box';
		frame.setAttribute('aria-hidden', 'true');
		Object.assign(frame.style, {
			left: `${ox + bounds.x * kx}px`,
			top: `${oy + bounds.y * ky}px`,
			width: `${bounds.width * kx}px`,
			height: `${bounds.height * ky}px`,
		});
		target.sheet.append(frame);
		this.#boxes.push(frame);
	}

	/** Drops the selection; true when there was one. */
	clear(): boolean {
		for (const box of this.#boxes.splice(0)) box.remove();
		this.#doc = undefined;
		if (!this.#targets.length) return false;
		this.#targets = [];
		return true;
	}
}
