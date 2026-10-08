import { visioArrangeCommands, type VisioArrangement } from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { RibbonCommand } from './ribbon-parts';

/** Arrangement controls submit one core-planned batch; the source transaction proves support. */
export class ViewerArrangement {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	#commands(state: ViewerState, action: VisioArrangement) {
		const page = state.document?.pages[state.pageIndex];
		if (
			!page ||
			!state.edit.sourceAvailable ||
			state.loading ||
			state.edit.busy ||
			state.selectedShapes.some((shape) => shape.pageId && shape.pageId !== page.id)
		)
			return undefined;
		return visioArrangeCommands(
			page,
			state.selectedShapes.map((shape) => shape.id),
			action,
		);
	}
	run(action: VisioArrangement): void {
		const commands = this.#commands(this.controller.state, action);
		if (commands)
			this.edit(
				() => this.controller.applyEdits(commands),
				action.type === 'align' ? 'Aligned selected shapes.' : 'Distributed selected shapes.',
			);
	}
	/** Returns whether Position can open for distribution even when rotation is unavailable. */
	render(state: ViewerState): boolean {
		const set = (id: string, available: boolean, minimum: number) => {
			const button = this.root.querySelector<RibbonCommand>(
				`[command="${id}"], [data-menu="${id}"]`,
			);
			if (!button) return;
			button.disabled = !available;
			const label = button.getAttribute('label') ?? '';
			button.title = available
				? `${label}: source locks and formulas may refuse this edit.`
				: `${label}: select ${minimum} or more local, unglued shapes in an editable .vsdx drawing.`;
		};
		const aligning = !!this.#commands(state, { type: 'align', edge: 'left' });
		for (const edge of ['left', 'center', 'right', 'top', 'middle', 'bottom'])
			set(`align-shapes-${edge}`, aligning, 2);
		set('align', aligning, 2);
		const distributing = !!this.#commands(state, { type: 'distribute', axis: 'horizontal' });
		set('distribute-horizontal', distributing, 3);
		set('distribute-vertical', distributing, 3);
		return distributing;
	}
}
