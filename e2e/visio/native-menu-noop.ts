import { expect, type Locator, type Page } from '@playwright/test';
type TestViewer = HTMLElement & {
	controller: {
		documentGeneration: number;
		state: {
			edit: { busy: boolean; dirty: boolean; canUndo: boolean; canRedo: boolean; error: unknown };
		};
	};
	exportVsdx(): { bytes: Uint8Array; dirty: boolean };
	applyEdits(
		edits: readonly { type: 'move-shape'; pageId: string; shapeId: string; x: number; y: number }[],
	): Promise<void>;
	undo(): Promise<void>;
	redo(): Promise<void>;
};
/** Native fully blocked commands must leave the source and an existing redo branch intact. */
export async function verifyBlockedMenuTransform(
	page: Page,
	viewer: Locator,
	target: Locator,
	source: Buffer,
	id: string,
	before: string,
) {
	const shape = viewer.locator(`[data-shape-id="${id}"]`);
	const generation = await viewer.evaluate(
		(element) => (element as TestViewer).controller.documentGeneration,
	);
	await target.click();
	await expect
		.poll(() => viewer.evaluate((element) => (element as TestViewer).controller.state.edit.busy))
		.toBe(false);
	await expect(viewer.locator('[data-status]')).toHaveText('No changes were made.');
	await expect(shape).toHaveAttribute('transform', before);
	const inspect = () =>
		viewer.evaluate((element) => {
			const host = element as TestViewer;
			const { busy, dirty, canUndo, canRedo, error } = host.controller.state.edit;
			return {
				busy,
				dirty,
				canUndo,
				canRedo,
				error,
				generation: host.controller.documentGeneration,
				bytes: Array.from(host.exportVsdx().bytes),
			};
		});
	expect(await inspect()).toEqual({
		busy: false,
		dirty: false,
		canUndo: false,
		canRedo: false,
		error: null,
		generation,
		bytes: Array.from(source),
	});
	await viewer.evaluate(async (element, id) => {
		const host = element as TestViewer;
		await host.applyEdits([{ type: 'move-shape', pageId: '0', shapeId: id, x: 1, y: 2 }]);
		await host.undo();
	}, id);
	const redoState = await inspect();
	expect(redoState.canRedo).toBe(true);
	expect(redoState.bytes).toEqual(Array.from(source));
	// Undo reloads the scene; restore selection and use the same actual menu command again.
	await shape.focus();
	await shape.press('Enter');
	await viewer
		.locator('[data-menu="position"]')
		.getByRole('button', { name: 'Position', exact: true })
		.click();
	await viewer
		.locator('[data-menu="rotate"]')
		.getByRole('menuitem', { name: 'Rotate Shapes', exact: true })
		.press('ArrowRight');
	await target.click();
	await expect
		.poll(() => viewer.evaluate((element) => (element as TestViewer).controller.state.edit.busy))
		.toBe(false);
	await expect(viewer.locator('[data-status]')).toHaveText('No changes were made.');
	expect(await inspect()).toEqual(redoState);
	await viewer.evaluate((element) => (element as TestViewer).redo());
	await expect(shape).not.toHaveAttribute('transform', before);
	await page.locator('#file').setInputFiles({
		name: 'blocked-copy.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: source,
	});
	await expect(page.locator('#file-name')).toHaveText('blocked-copy.vsdx');
	await expect(shape).toHaveAttribute('transform', before);
	const final = await inspect();
	expect(final.bytes).toEqual(Array.from(source));
	expect(final.dirty).toBe(false);
	expect(final.canUndo).toBe(false);
}
