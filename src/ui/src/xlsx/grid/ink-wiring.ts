// The grid's link to the Draw tab's surface (`ink-layer.ts`), which loads on first use. Pen
// strokes drawn while it loads are kept (`ink-early.ts`) and handed to it in order; a chunk that
// fails to load is tried again on the next use; a grid destroyed meanwhile builds nothing.
import type { GridInk, GridInkTool } from 'ooxml-core/xlsx/ui';
import type { GridView } from './grid-view';
import { EarlyInk } from './ink-early';
import type { InkLayer } from './ink-layer';

export interface GridInkWiring {
	ink: GridInk;
	/** For the grid's pointer handling: whether a press on the cells is ink (never a selection). */
	inking(event: PointerEvent): boolean;
	destroy(): void;
}

export function wireGridInk(view: GridView): GridInkWiring {
	let layer: InkLayer | undefined;
	let loading: Promise<InkLayer> | undefined;
	let tool: GridInkTool | undefined;
	let destroyed = false;
	const early = new EarlyInk(view);
	const report = (cause: unknown) =>
		view.ctx.toast(cause instanceof Error ? cause.message : String(cause), 'error');
	const load = () =>
		(loading ??= Promise.all([
			import('./ink-layer'),
			import('../../ink/tools').then((tools) => tools.loadTools()),
		]).then(
			([{ InkLayer }]) => {
				if (destroyed) throw new Error('The grid was destroyed.');
				const built = new InkLayer(view);
				layer = built;
				built.setTool(tool);
				for (const stroke of early.take()) built.stroke(stroke, stroke.pen);
				early.deliver = (stroke) => built.stroke(stroke, stroke.pen);
				return built;
			},
			(cause: unknown) => {
				loading = undefined;
				throw cause;
			},
		));
	const start = () => void load().catch((cause: unknown) => !destroyed && report(cause));
	return {
		ink: {
			setTool(next) {
				tool = next;
				if (layer) layer.setTool(next);
				else if (next) start();
			},
			hasSelection: () => layer?.hasSelection() ?? false,
			inkToShape: async () => (await load()).inkToShape(),
			replay: () =>
				void load().then(
					(built) => built.replay(),
					(cause: unknown) => report(cause),
				),
			replaying: () => layer?.replaying() ?? false,
		},
		inking(event) {
			if (tool === undefined || view.ctx.readOnly()) return false;
			if (!layer) {
				early.press(event, tool);
				// A load that failed is tried again with the next stroke.
				start();
			}
			return true;
		},
		destroy() {
			destroyed = true;
			early.destroy();
			layer?.destroy();
		},
	};
}
