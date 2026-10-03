import { commandRow, group, hint, tool } from './ribbon-parts.js';

/** Visio's View tab: Views, Show (grid and task panes) and Zoom. */
export function buildViewPanel(doc: Document, panel: HTMLElement): void {
	panel.append(
		commandRow(doc, 'View commands', [
			group(doc, 'Views', [
				tool(
					doc,
					'fullscreen',
					'Full Screen',
					'fullscreen',
					{ type: 'fullscreen' },
					{
						keys: ['F5', 'F5'],
						pressed: false,
					},
				),
			]),
			group(doc, 'Show', [
				tool(doc, 'grid', 'Grid', 'grid', { type: 'grid' }, { pressed: false }),
				tool(
					doc,
					'pages',
					'Pages pane',
					'visioPagesPane',
					{ type: 'pane', pane: 'pages' },
					{
						pressed: true,
						controls: 'page-rail',
					},
				),
				tool(
					doc,
					'inspector',
					'Inspector pane',
					'visioInspectorPane',
					{
						type: 'pane',
						pane: 'inspector',
					},
					{ pressed: true, controls: 'inspector-pane' },
				),
				tool(doc, 'notes', 'Review notes', 'message', { type: 'reveal', panel: 'notes' }),
			]),
			group(doc, 'Zoom', [
				tool(
					doc,
					'zoom-fit',
					'Fit to Window',
					'fitPage',
					{ type: 'zoom', mode: 'fit' },
					{
						keys: ['Control+Shift+W', 'Ctrl+Shift+W'],
					},
				),
				tool(doc, 'page-width', 'Page Width', 'pageWidth', { type: 'zoom', mode: 'width' }),
				tool(doc, 'actual-size', '100%', 'search', { type: 'zoom', mode: 'actual' }),
			]),
		]),
		hint(doc, ['Ctrl+wheel or + / − to zoom.', 'Ctrl+Page Up / Page Down changes page.']),
	);
}
