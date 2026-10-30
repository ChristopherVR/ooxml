import {
	visioOrgAddPosition,
	visioOrgDeletePosition,
	visioOrgHierarchy,
	visioOrgImportEdits,
	visioOrgPeople,
	visioOrgPosition,
	visioOrgReLayout,
	visioSelectionIsOnPage,
} from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController, type ViewerState } from './controller';
import type { VisioOrgChartAction } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';
import { DataFilePicker, readTable, readWorkbook } from './viewer-data-import';

type Command = VisioOrgChartAction['command'];

/** The org chart position the first selected shape is, on the displayed page. */
export function orgChartSelection(state: ViewerState): { pageId: string; id: string } | undefined {
	const page = state.document?.pages[state.pageIndex];
	const first = state.selectedShapes[0];
	if (!page || !first || !visioSelectionIsOnPage(first, page.id)) return undefined;
	return visioOrgPosition(page, first.id) ? { pageId: page.id, id: first.id } : undefined;
}

/**
 * The edits that delete the selected shape when it is an org chart position: Delete Position,
 * which removes its connectors and joins its subordinates to its manager. The Delete key uses
 * it for a lone selected position, as Visio's org chart does.
 */
export function orgChartDeleteSelection(
	state: ViewerState,
): { edits: ReturnType<typeof visioOrgDeletePosition>; message: string } | undefined {
	if (state.selectedShapes.length !== 1) return undefined;
	const selection = orgChartSelection(state);
	const page = state.document?.pages[state.pageIndex];
	if (!selection || !page) return undefined;
	const edits = visioOrgDeletePosition(page, selection.id);
	return edits ? { edits, message: 'Deleted the position and its connectors.' } : undefined;
}

/** The Org Chart tab's controller: one edit batch per command, so undo steps back one command. */
export class ViewerOrgChart {
	#picker: DataFilePicker | undefined;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	reason(state: ViewerState, command: Command): string | undefined {
		if (!canEditDiagram(state)) return 'Open a .vsdx file to edit the org chart.';
		if (state.loading || state.edit.busy) return 'Wait for the current operation to finish.';
		if (command === 'import') return undefined;
		const selection = orgChartSelection(state);
		if (!selection) return 'Select a position in an org chart.';
		const page = state.document!.pages[state.pageIndex]!;
		if (command === 'coworker' && !visioOrgHierarchy(page).parentOf(selection.id))
			return 'The top position has no manager to add a coworker under.';
		return undefined;
	}
	/** Org Chart > Import: a CSV or Excel sheet with Name, Title and Reports To columns. */
	#import(): void {
		this.#picker ??= (() => {
			const picker = new DataFilePicker(this.root);
			// The Data tab's own file input is found by this attribute; keep the two apart.
			delete picker.input.dataset.importData;
			picker.input.dataset.importOrg = '';
			return picker;
		})();
		const picker = this.#picker;
		this.edit(async () => {
			const picked = await picker.pick();
			if (!picked) return;
			const table = readTable(await readWorkbook(picked.file), {
				sheet: 0,
				range: '',
				header: true,
			});
			const people = visioOrgPeople(table);
			if (typeof people === 'string') throw new Error(people);
			const state = this.controller.state;
			const page = state.document?.pages[state.pageIndex];
			if (!page || this.reason(state, 'import') !== undefined) return;
			const plan = visioOrgImportEdits(page, people);
			if (typeof plan === 'string') throw new Error(plan);
			await this.controller.applyEdits(plan.edits);
			this.controller.selectShapes([{ id: plan.firstId, name: '', pageId: page.id }]);
		}, 'Imported the org chart.');
	}
	run(action: VisioOrgChartAction): void {
		const state = this.controller.state;
		if (this.reason(state, action.command) !== undefined) return;
		if (action.command === 'import') return this.#import();
		const page = state.document?.pages[state.pageIndex];
		const selection = orgChartSelection(state);
		if (!page || !selection) return;
		const { id } = selection;
		if (action.command === 'delete') {
			const plan = orgChartDeleteSelection(state);
			if (!plan?.edits) return;
			const edits = plan.edits;
			this.edit(async () => {
				await this.controller.applyEdits(edits);
				this.controller.selectShapes([]);
			}, plan.message);
		} else if (action.command === 'layout') {
			const edits = visioOrgReLayout(page);
			if (!edits) return;
			this.edit(() => this.controller.applyEdits(edits), 'Laid out the org chart.');
		} else {
			const kind = { subordinate: 'subordinate', coworker: 'peer', assistant: 'assistant' }[
				action.command
			] as 'subordinate' | 'peer' | 'assistant';
			const plan = visioOrgAddPosition(page, id, kind);
			if (!plan) return;
			this.edit(async () => {
				await this.controller.applyEdits(plan.edits);
				this.controller.selectShapes([{ id: plan.newId, name: '', pageId: page.id }]);
			}, `Added a ${action.command}.`);
		}
	}
	render(state: ViewerState): void {
		for (const [id, command] of [
			['org-subordinate', 'subordinate'],
			['org-coworker', 'coworker'],
			['org-assistant', 'assistant'],
			['org-delete', 'delete'],
			['org-layout', 'layout'],
			['org-import', 'import'],
			['org-import-data', 'import'],
		] as const) {
			// The Org Chart tab and the Data tab both hold an Import command.
			for (const item of this.root.querySelectorAll<RibbonCommand>(`[command="${id}"]`)) {
				const reason = this.reason(state, command);
				item.disabled = reason !== undefined;
				if (reason) item.title = reason;
			}
		}
	}
}
