// A pen stroke on Word's pages becomes floating ink placed where it was drawn, anchored to the
// body paragraph above it. Strokes drawn one after another, close in time and space, join the
// same drawing, as Word collects them, up to a size whose part is always read back whole.
import type { InkCoordinate, InkPen } from 'ooxml-core/drawingml';
import type { EditorView } from 'prosemirror-view';
import type { InkStrokeDetail } from '../ink/ink-canvas';
import { inkTools } from '../ink/tools';
import { bodyAnchor, scaleOf, sheetAt } from './draw-targets';
import { appendInk, insertInk } from './ink-commands';
import { inkOf } from './ink-node-view';

/** Ink drawn this close in time and space joins the previous drawing. */
const JOIN_MS = 2000;
const JOIN_PX = 48;

export interface StrokeHost {
	view(): EditorView | undefined;
	pages(): HTMLElement | undefined;
	pen(): InkPen;
	client(detail: { x: number; y: number }): InkCoordinate;
}

export class InkStrokes {
	readonly #host: StrokeHost;
	/** The ink the last stroke made, with its page (sheets are rebuilt by every layout). */
	#last: { pos: number; ink: string; page: number; at: number } | undefined;

	constructor(host: StrokeHost) {
		this.#host = host;
	}

	forget(): void {
		this.#last = undefined;
	}

	/** Adds a stroke; false when it cannot go on the page (no paragraph to anchor it). */
	add(detail: InkStrokeDetail): boolean {
		const view = this.#host.view();
		const first = detail.samples[0];
		if (!view || !first) return true;
		const sheet = sheetAt(this.#host.pages(), this.#host.client(first));
		if (!sheet) return true;
		const rect = sheet.getBoundingClientRect();
		const scale = scaleOf(sheet);
		const end = detail.samples.at(-1)?.time ?? 0;
		const samples = detail.samples.map((sample) => {
			const client = this.#host.client(sample);
			return {
				x: (client.x - rect.left) / scale,
				y: (client.y - rect.top) / scale,
				pressure: sample.pressure,
				time: Date.now() - end + sample.time,
			};
		});
		const tools = inkTools();
		const stroke = tools.inkStrokeFromSamples(samples, { ...this.#host.pen(), kind: detail.tool });
		if (!stroke) return true;
		const now = Date.now();
		const page = Number(sheet.dataset.pageIndex ?? -1);
		const pos = this.#join(view, stroke, page, now);
		if (pos !== undefined) return this.#remember(view, pos, page, now);
		const blockId = bodyAnchor(view, sheet, this.#host.client(first).y);
		const at = blockId ? insertInk(view, [stroke], { blockId }) : undefined;
		return at === undefined ? false : this.#remember(view, at, page, now);
	}

	/** Adds the stroke to the last drawing when it is near it in time and space. */
	#join(
		view: EditorView,
		stroke: NonNullable<ReturnType<ReturnType<typeof inkTools>['inkStrokeFromSamples']>>,
		page: number,
		now: number,
	): number | undefined {
		const last = this.#last;
		const node = last && view.state.doc.nodeAt(last.pos);
		const ink = node && node.attrs.ink === last?.ink ? inkOf(node) : undefined;
		const bounds = inkTools().inkStrokesBounds([stroke]);
		if (!last || !node || !ink || !bounds || now - last.at >= JOIN_MS || last.page !== page)
			return undefined;
		const placement = JSON.parse(String(node.attrs.placement)) as {
			offsetXPx?: number;
			offsetYPx?: number;
		};
		const x = placement.offsetXPx ?? 0;
		const y = placement.offsetYPx ?? 0;
		const near =
			bounds.x < x + Number(node.attrs.widthPx) + JOIN_PX &&
			bounds.x + bounds.width > x - JOIN_PX &&
			bounds.y < y + Number(node.attrs.heightPx) + JOIN_PX &&
			bounds.y + bounds.height > y - JOIN_PX;
		return near &&
			inkTools().canJoinInk(ink.strokes, [stroke]) &&
			appendInk(view, last.pos, [stroke])
			? last.pos
			: undefined;
	}

	#remember(view: EditorView, pos: number, page: number, at: number): true {
		const ink = view.state.doc.nodeAt(pos)?.attrs.ink;
		this.#last = typeof ink === 'string' ? { pos, ink, page, at } : undefined;
		return true;
	}
}
