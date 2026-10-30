import { command, commandRow, group } from './ribbon-parts';

/**
 * Visio's contextual Org Chart tab, as far as positions go: Insert (Subordinate, Coworker,
 * Assistant), Position (Delete Position), Import and Arrange (Layout). It shows while an org chart
 * position is selected (`viewer-contextual-tabs.ts`); `viewer-org-chart.ts` keeps its state.
 * Import builds a chart from a CSV or Excel sheet (Name, Title and Reports To columns); a
 * directory is not read. The Data tab offers the same command, because it needs no selection.
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
			group(doc, 'Import', [
				command(doc, {
					id: 'org-import',
					label: 'Import Org Chart',
					icon: 'visioData',
					action: { type: 'org-chart', command: 'import' },
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
