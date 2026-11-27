import type { EditorContext } from 'ooxml-core/xlsx/ui';
import type { createChartExPane } from './chart-ex-pane';
import type { createChartFormatPane } from './chart-format-pane';

/** Mount the chart panes in keyboard order as their commands are first used. */
export function registerChartPanes(
	ctx: EditorContext,
	workspace: HTMLElement,
	smartArt: HTMLElement,
) {
	let format: ReturnType<typeof createChartFormatPane> | undefined;
	let extended: ReturnType<typeof createChartExPane> | undefined;
	const loadFormat = async () => {
		if (format) return format;
		const { createChartFormatPane } = await import('./chart-format-pane');
		if (!format) {
			format = createChartFormatPane(ctx);
			workspace.after(format.element);
		}
		return format;
	};
	ctx.dialogs.register('format-chart-area', async (_ctx, arg) =>
		(await loadFormat()).openArea(arg),
	);
	ctx.dialogs.register('format-chart-series', async (_ctx, arg) =>
		(await loadFormat()).openSeries(arg),
	);
	ctx.dialogs.register('format-chart-selection', async (_ctx, arg) =>
		(await loadFormat()).openTarget(arg),
	);
	ctx.dialogs.register('chart-ex', async () => {
		if (!extended) {
			const { createChartExPane } = await import('./chart-ex-pane');
			if (!extended) {
				extended = createChartExPane(ctx);
				smartArt.before(extended.element);
			}
		}
		extended.open();
	});
	return {
		elements: () => [format?.element, extended?.element].filter((element) => element !== undefined),
		refresh() {
			format?.refresh();
			extended?.refresh();
		},
		relocalize() {
			format?.relocalize();
			extended?.relocalize();
		},
	};
}
