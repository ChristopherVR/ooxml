import { check, command, commandRow, group, menu, stack } from './ribbon-parts';

const EDITING_AIDS = 'Needs interactive shape editing aids.';
const WINDOWS = 'The viewer shows one drawing window.';
const MACROS = 'Macros and add-ons never run in this viewer.';

/**
 * Microsoft Visio's View tab: Views, Show, Zoom, Visual Aids, Window and Macros. Ruler, Grid,
 * Task Panes, Presentation Mode, Full Screen and every Zoom command work; the rest is shown disabled with a reason.
 */
export function buildViewPanel(doc: Document, panel: HTMLElement): void {
	panel.append(
		commandRow(doc, 'View commands', [
			group(doc, 'Views', [
				command(doc, {
					id: 'presentation',
					label: 'Presentation Mode',
					icon: 'visioPresentation',
					action: { type: 'presentation' },
					keys: ['F5', 'F5'],
				}),
				command(doc, {
					id: 'fullscreen',
					label: 'Full Screen',
					icon: 'fullscreen',
					action: { type: 'fullscreen' },
					pressed: false,
				}),
			]),
			group(doc, 'Show', [
				stack(doc, [
					check(doc, { id: 'ruler', label: 'Ruler', action: { type: 'ruler' } }),
					check(doc, {
						id: 'page-breaks',
						label: 'Page Breaks',
						action: { type: 'page-setup', command: { op: 'page-breaks' } },
					}),
					check(doc, { id: 'grid', label: 'Grid', action: { type: 'grid' } }),
				]),
				stack(doc, [
					check(doc, {
						id: 'guides',
						label: 'Guides',
						unsupported: 'Needs guide shapes from the core.',
					}),
					menu(doc, {
						id: 'task-panes',
						label: 'Task Panes',
						icon: 'visioInspectorPane',
						size: 'small',
						items: [
							{
								id: 'shapes',
								label: 'Shapes',
								action: { type: 'pane', pane: 'shapes' },
								checked: true,
							},
							{
								id: 'inspector',
								label: 'Inspector',
								action: { type: 'pane', pane: 'inspector' },
								checked: true,
							},
							{
								id: 'shape-data',
								label: 'Shape Data',
								action: { type: 'reveal', panel: 'selection' },
							},
							{ id: 'layers-pane', label: 'Layers', action: { type: 'reveal', panel: 'layers' } },
							{
								id: 'notes',
								label: 'Compatibility Notes',
								action: { type: 'reveal', panel: 'notes' },
							},
							{ id: 'pan-zoom', label: 'Pan & Zoom', action: { type: 'panZoom' }, checked: false },
							{
								id: 'size-position',
								label: 'Size & Position',
								action: { type: 'sizePosition' },
								checked: false,
							},
							{
								id: 'drawing-explorer',
								label: 'Drawing Explorer',
								unsupported: 'Needs a document tree view.',
							},
						],
					}),
				]),
			]),
			group(doc, 'Zoom', [
				menu(doc, {
					id: 'zoom',
					label: 'Zoom',
					icon: 'zoomIn',
					items: [400, 200, 150, 100, 75, 50].map((percent) => ({
						id: `zoom-${percent}`,
						label: `${percent}%`,
						action: { type: 'zoomTo', percent },
					})),
				}),
				command(doc, {
					id: 'zoom-fit',
					label: 'Fit to Window',
					icon: 'fitPage',
					action: { type: 'zoom', mode: 'fit' },
					keys: ['Control+Shift+W', 'Ctrl+Shift+W'],
				}),
				command(doc, {
					id: 'page-width',
					label: 'Page Width',
					icon: 'pageWidth',
					action: { type: 'zoom', mode: 'width' },
				}),
			]),
			group(
				doc,
				'Visual Aids',
				[
					stack(doc, [
						check(doc, {
							id: 'auto-connect',
							label: 'AutoConnect',
							unsupported:
								'Hover arrows that add a connected copy need shape duplication with glue in one step, which the core does not offer yet.',
						}),
						check(doc, { id: 'dynamic-grid', label: 'Dynamic Grid', unsupported: EDITING_AIDS }),
						check(doc, {
							id: 'connection-points',
							label: 'Connection Points',
							action: { type: 'connection-points' },
							checked: true,
						}),
					]),
				],
				{ launcher: EDITING_AIDS },
			),
			group(doc, 'Window', [
				command(doc, { id: 'new-window', label: 'New Window', icon: 'copy', unsupported: WINDOWS }),
				stack(doc, [
					command(doc, {
						id: 'arrange-all',
						label: 'Arrange All',
						icon: 'grid',
						size: 'small',
						unsupported: WINDOWS,
					}),
					command(doc, {
						id: 'cascade',
						label: 'Cascade',
						icon: 'bringToFront',
						size: 'small',
						unsupported: WINDOWS,
					}),
				]),
				command(doc, {
					id: 'switch-windows',
					label: 'Switch Windows',
					icon: 'visioPagesPane',
					unsupported: WINDOWS,
				}),
			]),
			group(doc, 'Macros', [
				command(doc, { id: 'macros', label: 'Macros', icon: 'visioMacros', unsupported: MACROS }),
				command(doc, { id: 'add-ons', label: 'Add-Ons', icon: 'settings', unsupported: MACROS }),
			]),
		]),
	);
}
