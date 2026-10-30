import type { VisioAction, VisioShape } from 'ooxml-core/visio';
import { editErrorMessage, isEditCancellation } from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController } from './controller';

const ITEM = 'data-shape-action';
const COMMAND = 'ctx-action:';
const listening = new WeakSet<HTMLElement>();

/** Visio orders a shape's actions by SortKey when any row has one, otherwise by row. */
export function orderedActions(actions: readonly VisioAction[]): VisioAction[] {
	const keyed = actions.some((action) => action.sortKey);
	return actions
		.map((action, index) => ({ action, index }))
		.sort((a, b) =>
			keyed
				? (a.action.sortKey ?? '').localeCompare(b.action.sortKey ?? '') || a.index - b.index
				: a.index - b.index,
		)
		.map((item) => item.action);
}

/** Why an action cannot run now, or undefined when it can. */
export function actionBlock(action: VisioAction, editable: boolean): string | undefined {
	if (action.disabled) return 'The shape has disabled this action.';
	if (action.run.kind === 'unsupported') return action.run.reason;
	return editable ? undefined : 'The drawing cannot be edited.';
}

function selectedShape(controller: Pick<ViewerController, 'state'>): VisioShape | undefined {
	const { state } = controller;
	const page = state.document?.pages[state.pageIndex];
	if (!page || state.selectedShapes.length !== 1) return undefined;
	const id = state.selectedShape?.id;
	return page.shapes.find((shape) => shape.id === id);
}

/**
 * Adds the selected shape's Actions rows to its context menu, as Visio lists them below the
 * built-in commands: checked and disabled states are kept, and an action this editor cannot run
 * is shown disabled with the reason as its tip. Called each time the shape menu opens.
 */
export function syncShapeActions(
	menu: HTMLElement,
	target: 'shape' | 'page',
	controller: Pick<ViewerController, 'state' | 'applyEdits'>,
	announce: (message: string) => void,
): void {
	for (const node of [...menu.querySelectorAll(`:scope > [${ITEM}]`)]) node.remove();
	if (target !== 'shape') return;
	const shape = selectedShape(controller);
	const actions = orderedActions((shape?.actions ?? []).filter((action) => !action.invisible));
	if (!shape || !actions.length) return;
	const doc = menu.ownerDocument;
	const editable = canEditDiagram(controller.state);
	const separator = (): HTMLElement => {
		const rule = doc.createElement('office-ui-menu-separator');
		rule.setAttribute(ITEM, '');
		return rule;
	};
	menu.append(separator());
	for (const action of actions) {
		if (action.beginGroup && menu.lastElementChild?.localName !== 'office-ui-menu-separator')
			menu.append(separator());
		const item = doc.createElement('office-ui-menu-item');
		item.setAttribute(ITEM, action.row);
		item.setAttribute('command', `${COMMAND}${action.row}`);
		item.setAttribute('label', action.menu || action.name || action.row);
		if (action.checked) item.setAttribute('checked', 'true');
		const block = actionBlock(action, editable);
		if (block) {
			item.setAttribute('disabled', '');
			item.setAttribute('title', block);
		}
		menu.append(item);
	}
	if (listening.has(menu)) return;
	listening.add(menu);
	menu.addEventListener('office-command', (event) => {
		const command = (event as CustomEvent<{ command: string }>).detail?.command ?? '';
		if (!command.startsWith(COMMAND)) return;
		const owner = selectedShape(controller);
		const page = controller.state.document?.pages[controller.state.pageIndex];
		if (!owner || !page) return;
		const row = command.slice(COMMAND.length);
		const action = owner.actions?.find((item) => item.row === row);
		if (!action || actionBlock(action, canEditDiagram(controller.state))) return;
		controller
			.applyEdits([{ type: 'run-action', pageId: page.id, shapeId: owner.id, row }])
			.then(() => announce(`${action.menu || 'Action'} done.`))
			.catch((error: unknown) => {
				if (!isEditCancellation(error)) announce(editErrorMessage(error));
			});
	});
}
