import { command, commandRow, group } from './ribbon-parts';

/**
 * Visio's contextual Org Chart tab, as far as positions go: Insert (Subordinate, Coworker,
 * Assistant), Position (Delete Position) and Arrange (Layout). It shows while an org chart
 * position is selected (`viewer-contextual-tabs.ts`); `viewer-org-chart.ts` keeps its state.
 * Importing people from a file or directory is not offered.
 */
export function buildOrgChartPanel(doc: Document, panel: HTMLElement): void {
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Org Chart commands', [
			group(doc, 'Insert', [
				command(doc, {
					id: 'org-subordinate',
					label: 'Subordinate',
					icon: 'rectangle',
					action: { type: 'org-chart', command: 'subordinate' },
				}),
				command(doc, {
					id: 'org-coworker',
					label: 'Coworker',
					icon: 'rectangle',
					action: { type: 'org-chart', command: 'coworker' },
				}),
				command(doc, {
					id: 'org-assistant',
					label: 'Assistant',
					icon: 'rectangle',
					action: { type: 'org-chart', command: 'assistant' },
				}),
			]),
			group(doc, 'Position', [
				command(doc, {
					id: 'org-delete',
					label: 'Delete Position',
					icon: 'delete',
					action: { type: 'org-chart', command: 'delete' },
				}),
			]),
			group(doc, 'Arrange', [
				command(doc, {
					id: 'org-layout',
					label: 'Layout',
					icon: 'position',
					action: { type: 'org-chart', command: 'layout' },
				}),
			]),
		]),
	);
}
