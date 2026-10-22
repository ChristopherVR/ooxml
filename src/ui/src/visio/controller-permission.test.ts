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

it('rejects an in-flight local edit when the host revokes permission', async () => {
	let finish!: (result: import('./worker-editor').EditTransactionResult) => void;
	const edited = new Promise<import('./worker-editor').EditTransactionResult>((resolve) => {
		finish = resolve;
	});
	const controller = new ViewerController(
		async () => structuredClone(demoDocument),
		() => {},
		() => edited,
	);
	await controller.load(new Uint8Array([1]));
	const operation = controller.applyEdits([]);
	expect(controller.state.edit.busy).toBe(true);
	controller.canEdit = false;
	finish({
		bytes: new Uint8Array([2]),
		document: structuredClone(demoDocument),
		changedParts: ['page.xml'],
		diagnostics: [],
	});
	await expect(operation).rejects.toMatchObject({ name: 'AbortError' });
	expect(controller.exportVsdx().bytes).toEqual(new Uint8Array([1]));
	expect(controller.state.edit.dirty).toBe(false);
	expect(controller.canEdit).toBe(false);
	controller.destroy();
});

it('readOnly, canEdit and a viewer role each restrict and none lifts another', async () => {
	const controller = new ViewerController(
		async () => structuredClone(demoDocument),
		() => {},
		vi.fn() as never,
	);
	await controller.load(new Uint8Array([1]));
	expect(controller.readOnly).toBe(false);
	controller.readOnly = true;
	expect(controller.canEdit).toBe(false);
	controller.canEdit = true;
	expect(controller.canEdit).toBe(false);
	controller.readOnly = false;
	expect(controller.canEdit).toBe(true);
	controller.setRoleReadOnly(true);
	controller.canEdit = true;
	controller.readOnly = false;
	expect(controller.readOnly).toBe(true);
	controller.setRoleReadOnly(false);
	expect(controller.canEdit).toBe(true);
	controller.canEdit = false;
	controller.setRoleReadOnly(true);
	controller.setRoleReadOnly(false);
	expect(controller.canEdit).toBe(false);
	controller.destroy();
});
