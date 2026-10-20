import { expect, it, vi } from 'vitest';
import { ViewerController } from './controller';
import { demoDocument } from 'ooxml-core/visio/ui';
import { diagramToolbarProps } from './building-blocks';

it('keeps navigation and selection available while rejecting direct mutations', async () => {
	const edit = vi.fn();
	const controller = new ViewerController(
		async () => structuredClone(demoDocument),
		() => {},
		edit,
	);
	await controller.load(new Uint8Array([1]));
	controller.canEdit = false;
	controller.setZoom(2);
	controller.setPage(1);
	controller.setSearchQuery('a');
	controller.selectAll();
	expect(controller.state.zoom).toBe(2);
	expect(controller.state.canEdit).toBe(false);
	await expect(controller.applyEdits([])).rejects.toThrow('read-only');
	await expect(controller.undo()).rejects.toThrow('read-only');
	await expect(controller.redo()).rejects.toThrow('read-only');
	expect(edit).not.toHaveBeenCalled();
	expect(controller.exportVsdx().bytes).toEqual(new Uint8Array([1]));
	const props = diagramToolbarProps(controller, () => {});
	expect(props.canEdit).toBe(false);
	expect(props.canUndo).toBe(false);
	controller.destroy();
});
