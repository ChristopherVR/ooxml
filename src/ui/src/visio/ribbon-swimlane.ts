import { VISIO_SWIMLANE_STYLES } from 'ooxml-core/visio';
import { check, command, commandRow, group, menu, stack } from './ribbon-parts';

const DIRECTION =
	'The direction is chosen when the flowchart is created (File > New > Cross-Functional Flowchart); an existing flowchart cannot be turned yet.';

/**
 * Visio's contextual Cross-Functional Flowchart tab (Microsoft's "Create a cross-functional
 * flowchart"): Insert (Swimlane, Separator), Arrange (Direction, Rotate Lane Labels) and Design
 * (Style, Show Title Bar, Show Separator Bar). It shows while a flowchart's frame, a lane or a
 * shape in a lane is selected (`viewer-contextual-tabs.ts`); `viewer-swimlanes.ts` keeps its
 * state and availability.
 */
export function buildSwimlanePanel(doc: Document, panel: HTMLElement): void {
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Cross-Functional Flowchart commands', [
			group(doc, 'Insert', [
				menu(doc, {
					id: 'swimlane-insert',
					label: 'Swimlane',
					icon: 'rectangle',
					items: [
						{
							id: 'swimlane-insert-before',
							label: 'Insert Swimlane Before',
							action: { type: 'swimlane', command: 'insert-before' },
						},
						{
							id: 'swimlane-insert-after',
							label: 'Insert Swimlane After',
							action: { type: 'swimlane', command: 'insert-after' },
						},
						{
							id: 'swimlane-delete',
							label: 'Delete Swimlane',
							action: { type: 'swimlane', command: 'delete' },
						},
					],
				}),
				command(doc, {
					id: 'swimlane-separator',
					label: 'Separator',
					icon: 'line',
					action: { type: 'swimlane', command: 'separator' },
				}),
			]),
			group(doc, 'Arrange', [
				menu(doc, {
					id: 'swimlane-direction',
					label: 'Direction',
					icon: 'position',
					items: [
						{ id: 'swimlane-horizontal', label: 'Horizontal', unsupported: DIRECTION },
						{ id: 'swimlane-vertical', label: 'Vertical', unsupported: DIRECTION },
					],
				}),
				command(doc, {
					id: 'swimlane-rotate-labels',
					label: 'Rotate Lane Labels',
					icon: 'visioTextBlock',
					action: { type: 'swimlane', command: 'rotate-labels' },
					pressed: false,
				}),
			]),
			group(doc, 'Design', [
				menu(doc, {
					id: 'swimlane-style',
					label: 'Style',
					icon: 'effects',
					items: VISIO_SWIMLANE_STYLES.map((style) => ({
						id: `swimlane-style-${style}`,
						label: `Style ${style}`,
						action: { type: 'swimlane', command: 'style', style },
					})),
				}),
				stack(doc, [
					check(doc, {
						id: 'swimlane-title',
						label: 'Show Title Bar',
						action: { type: 'swimlane', command: 'title' },
						checked: true,
					}),
					check(doc, {
						id: 'swimlane-separators',
						label: 'Show Separator Bar',
						action: { type: 'swimlane', command: 'separators' },
						checked: true,
					}),
				]),
			]),
		]),
	);
}
