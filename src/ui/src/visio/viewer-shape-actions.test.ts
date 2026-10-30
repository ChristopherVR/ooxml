import type { VisioAction } from 'ooxml-core/visio';
import { expect, it, vi } from 'vitest';
import { actionBlock, orderedActions, syncShapeActions } from './viewer-shape-actions';

const action = (row: string, extra: Partial<VisioAction> = {}): VisioAction => ({
	row,
	menu: `Action ${row}`,
	checked: false,
	disabled: false,
	invisible: false,
	beginGroup: false,
	run: { kind: 'setf', targets: ['User.Mode'] },
	...extra,
});

function setup(actions: VisioAction[], sourceAvailable = true) {
	const menu = document.createElement('div');
	menu.append(document.createElement('office-ui-menu-item'));
	const applyEdits = vi.fn(async () => {});
	const controller = {
		applyEdits,
		state: {
			canEdit: true,
			edit: { sourceAvailable },
			pageIndex: 0,
			selectedShape: { id: '4', name: 'Box' },
			selectedShapes: [{ id: '4', name: 'Box' }],
			document: { pages: [{ id: '1', shapes: [{ id: '4', actions }] }] },
		},
	};
	const announce = vi.fn();
	return { menu, controller, announce, applyEdits };
}

it('orders by SortKey when a row has one, otherwise by row', () => {
	const rows = [action('a', { sortKey: 'b' }), action('b', { sortKey: 'a' }), action('c')];
	expect(orderedActions(rows).map((item) => item.row)).toEqual(['c', 'b', 'a']);
	expect(orderedActions([action('x'), action('y')]).map((item) => item.row)).toEqual(['x', 'y']);
});

it('explains a disabled, unsupported or unavailable action', () => {
	expect(actionBlock(action('a'), true)).toBeUndefined();
	expect(actionBlock(action('a', { disabled: true }), true)).toMatch(/disabled/);
	expect(
		actionBlock(
			action('a', { run: { kind: 'unsupported', reason: 'OPENFILE is not run here.' } }),
			true,
		),
	).toBe('OPENFILE is not run here.');
	expect(actionBlock(action('a'), false)).toMatch(/cannot be edited/);
});

it('lists visible actions with their state and runs the chosen one', async () => {
	const ui = setup([
		action('1', { checked: true, beginGroup: true }),
		action('2', { invisible: true }),
		action('3', { disabled: true }),
		action('4', { run: { kind: 'unsupported', reason: 'RUNADDON is not run here.' } }),
	]);
	syncShapeActions(ui.menu, 'shape', ui.controller as never, ui.announce);
	const items = [...ui.menu.querySelectorAll('[data-shape-action]')].filter(
		(node) => node.localName === 'office-ui-menu-item',
	);
	expect(items.map((node) => node.getAttribute('data-shape-action'))).toEqual(['1', '3', '4']);
	expect(items[0]!.getAttribute('checked')).toBe('true');
	expect(items[1]!.hasAttribute('disabled')).toBe(true);
	expect(items[2]!.getAttribute('title')).toBe('RUNADDON is not run here.');
	// Opening the menu again replaces the entries instead of adding to them.
	syncShapeActions(ui.menu, 'shape', ui.controller as never, ui.announce);
	expect(ui.menu.querySelectorAll('office-ui-menu-item[data-shape-action]')).toHaveLength(3);
	ui.menu.dispatchEvent(
		new CustomEvent('office-command', { detail: { command: 'ctx-action:1' }, bubbles: true }),
	);
	await vi.waitFor(() =>
		expect(ui.applyEdits).toHaveBeenCalledWith([
			{ type: 'run-action', pageId: '1', shapeId: '4', row: '1' },
		]),
	);
	// A disabled row never runs, even when its command is sent.
	ui.menu.dispatchEvent(
		new CustomEvent('office-command', { detail: { command: 'ctx-action:3' }, bubbles: true }),
	);
	expect(ui.applyEdits).toHaveBeenCalledTimes(1);
});

it('adds nothing to the page menu or for a shape without actions', () => {
	const ui = setup([]);
	syncShapeActions(ui.menu, 'shape', ui.controller as never, ui.announce);
	expect(ui.menu.querySelectorAll('[data-shape-action]')).toHaveLength(0);
	const withActions = setup([action('1')]);
	syncShapeActions(withActions.menu, 'page', withActions.controller as never, withActions.announce);
	expect(withActions.menu.querySelectorAll('[data-shape-action]')).toHaveLength(0);
});
