import { expect, it, vi } from 'vitest';
import {
	captureVisioClipboard,
	createVsdx,
	editVsdx,
	parseVsdx,
	serializeVisioClipboard,
} from 'ooxml-core/visio';
import { ViewerController } from './controller';
import type { ViewerTextReplaceToken } from './replacement-token';
import type { CancellableEditor } from './worker-editor';

const edit: CancellableEditor = async (bytes, commands) => {
	const result = await editVsdx(bytes, commands);
	return { ...result, document: await parseVsdx(result.bytes) };
};
async function setup(service: CancellableEditor = edit) {
	const pages = await editVsdx(await createVsdx(), [
		{ type: 'insert-page', pageId: '1', afterPageId: '0', name: 'Other' },
	]);
	const source = await editVsdx(pages.bytes, [
		{
			type: 'create-rectangle',
			pageId: '0',
			shapeId: '1',
			x: 2,
			y: 2,
			width: 1,
			height: 1,
			text: 'aa aa\n',
		},
		{
			type: 'create-rectangle',
			pageId: '0',
			shapeId: '2',
			x: 4,
			y: 2,
			width: 1,
			height: 1,
			text: 'aa',
		},
		{
			type: 'create-rectangle',
			pageId: '1',
			shapeId: '3',
			x: 2,
			y: 2,
			width: 1,
			height: 1,
			text: '😀aa',
		},
	]);
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		service,
		async (bytes, page, ids) =>
			serializeVisioClipboard(await captureVisioClipboard(bytes, page, ids)),
	);
	await controller.load(source.bytes);
	controller.selectAll();
	return controller;
}
function deferred() {
	let release!: () => void;
	const waiting = new Promise<void>((resolve) => {
		release = resolve;
	});
	return { waiting, release };
}
it('cancels only an applied replacement owner, including after newer selection intent', async () => {
	const gate = deferred();
	const c = await setup(async (bytes, commands) => {
		await gate.waiting;
		return edit(bytes, commands);
	});
	const before = c.exportVsdx().bytes;
	const unstarted = c.captureTextReplaceToken('all-pages');
	const token = c.captureTextReplaceToken('all-pages');
	const plan = c.planTextReplacement(token, { query: 'aa', replacement: 'X', mode: 'all' });
	const pending = c.applyTextReplacePlan(plan, token).catch((error: unknown) => error);
	expect(c.state.edit.busy).toBe(true);
	c.cancelTextReplace(unstarted);
	c.cancelTextReplace(Object.freeze({}) as ViewerTextReplaceToken);
	expect(c.state.edit.busy).toBe(true);
	c.clearSelection();
	expect(c.isTextReplaceTokenCurrent(token)).toBe(false);
	c.cancelTextReplace(token);
	expect(c.state.edit.busy).toBe(false);
	gate.release();
	expect(await pending).toMatchObject({ name: 'AbortError' });
	expect(c.exportVsdx().bytes).toEqual(before);
	expect(c.state.edit.canUndo).toBe(false);
	c.destroy();
	c.cancelTextReplace(token); // Adapter disposal after controller destruction is a safe no-op.
});

it('a completed replacement token cannot cancel a subsequent unrelated source edit', async () => {
	const gate = deferred();
	let waiting = false;
	const c = await setup(async (bytes, commands) => {
		if (waiting) await gate.waiting;
		return edit(bytes, commands);
	});
	const token = c.captureTextReplaceToken('all-pages');
	const plan = c.planTextReplacement(token, { query: 'aa', replacement: 'X', mode: 'all' });
	const fresh = await c.applyTextReplacePlan(plan, token);
	waiting = true;
	const unrelated = c.applyEdits([
		{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'Unrelated' },
	]);
	c.cancelTextReplace(token);
	c.cancelTextReplace(fresh);
	expect(c.state.edit.busy).toBe(true);
	gate.release();
	await unrelated;
	expect(c.state.document!.pages[0]!.shapes[0]!.text.plainText).toBe('Unrelated');
	c.destroy();
});
it('retains ordered selection scope through owned navigation without accepting old tokens', async () => {
	const c = await setup();
	c.selectShapes([...c.state.selectedShapes].reverse());
	let token = c.captureTextReplaceToken('selection');
	const old = token;
	const matches = c.getTextReplaceOccurrences(token, 'aa');
	expect(matches.map((match) => match.shapeId)).toEqual(['2', '1', '1']);
	token = c.selectTextReplaceOccurrence(token, matches[1]!, 'aa');
	expect(c.state.selectedShapes.map((shape) => shape.id)).toEqual(['1']);
	expect(c.getTextReplaceOccurrences(token, 'aa')).toEqual(matches);
	expect(c.isTextReplaceTokenCurrent(old)).toBe(false);
	expect(Object.isFrozen(token) && Object.keys(token).length === 0).toBe(true);
	c.destroy();
});
it('accepts cross-page current replacement and next selection atomically with exact history', async () => {
	const c = await setup();
	const before = c.exportVsdx().bytes;
	const prior = c.state.selectedShapes;
	const token = c.captureTextReplaceToken('all-pages');
	const matches = c.getTextReplaceOccurrences(token, 'aa');
	const plan = c.planTextReplacement(token, {
		query: 'aa',
		replacement: '$&',
		mode: 'current',
		current: matches[2]!,
	});
	expect(plan.nextOccurrence).toEqual({ pageId: '1', shapeId: '3', start: 2, end: 4 });
	const primary = vi.fn();
	c.onEvent((name, detail) => {
		if (name === 'shape-select') primary(detail);
	});
	const publish = vi.fn();
	c.subscribe((state) => {
		if (state.document?.pages[0]!.shapes[1]!.text.plainText === '$&') {
			expect(state.pageIndex).toBe(1);
			expect(state.selectedShape?.id).toBe('3');
			publish();
		}
	});
	const next = await c.applyTextReplacePlan(plan, token);
	expect(publish).toHaveBeenCalled();
	expect(primary).toHaveBeenCalledOnce();
	expect(c.isTextReplaceTokenCurrent(next)).toBe(true);
	expect(c.getTextReplaceOccurrences(next, 'aa')).toHaveLength(3);
	const changed = c.exportVsdx().bytes;
	await c.undo();
	expect(c.exportVsdx().bytes).toEqual(before);
	expect(c.state.pageIndex).toBe(0);
	expect(c.state.selectedShapes).toEqual(prior);
	await c.redo();
	expect(c.exportVsdx().bytes).toEqual(changed);
	expect(c.state.pageIndex).toBe(1);
	expect(c.state.selectedShape?.id).toBe('3');
	c.destroy();
});
it('replaces all scoped targets as one edit, preserving literal trailing LF and exact undo', async () => {
	const c = await setup();
	const before = c.exportVsdx().bytes;
	const token = c.captureTextReplaceToken('all-pages');
	const plan = c.planTextReplacement(token, { query: 'aa', replacement: 'X\n', mode: 'all' });
	expect(plan.replacementCount).toBe(4);
	await c.applyTextReplacePlan(plan, token);
	expect(c.state.document!.pages[0]!.shapes[0]!.text.plainText).toBe('X\n X\n\n');
	expect(c.state.document!.pages[1]!.shapes[0]!.text.plainText).toBe('😀X\n');
	await c.undo();
	expect(c.exportVsdx().bytes).toEqual(before);
	c.destroy();
});
it('keeps no-op source/history exact while navigating to the next occurrence', async () => {
	const c = await setup();
	const before = c.exportVsdx().bytes;
	const token = c.captureTextReplaceToken('all-pages');
	const matches = c.getTextReplaceOccurrences(token, 'aa');
	const plan = c.planTextReplacement(token, {
		query: 'aa',
		replacement: 'aa',
		mode: 'current',
		current: matches[2]!,
	});
	const next = await c.applyTextReplacePlan(plan, token);
	expect(c.exportVsdx().bytes).toEqual(before);
	expect(c.state.edit.canUndo).toBe(false);
	expect(c.state.selectedShape?.id).toBe('3');
	expect(c.state.pageIndex).toBe(1);
	expect(c.isTextReplaceTokenCurrent(next)).toBe(true);
	c.destroy();
});
for (const reason of ['selection', 'page', 'cancel', 'model', 'destroy'] as const) {
	it(`rejects ${reason} changes during editor await without accepting source/history`, async () => {
		const gate = deferred();
		const c = await setup(async (bytes, commands) => {
			await gate.waiting;
			return edit(bytes, commands);
		});
		const before = c.exportVsdx().bytes;
		const token = c.captureTextReplaceToken('all-pages');
		const plan = c.planTextReplacement(token, { query: 'aa', replacement: 'X', mode: 'all' });
		const result = c.applyTextReplacePlan(plan, token);
		const failed = expect(result).rejects.toMatchObject({ name: 'AbortError' });
		if (reason === 'selection') {
			c.clearSelection();
			c.selectAll();
		}
		if (reason === 'page') {
			c.setPage(1);
			c.setPage(0);
		}
		if (reason === 'cancel') c.cancelEdit();
		if (reason === 'model') c.setDocument(c.state.document);
		if (reason === 'destroy') c.destroy();
		gate.release();
		await failed;
		if (reason === 'selection' || reason === 'page' || reason === 'cancel') {
			expect(c.exportVsdx().bytes).toEqual(before);
			expect(c.state.edit.canUndo).toBe(false);
		}
		if (reason !== 'destroy') c.destroy();
	});
}
it('rejects forged/foreign tokens and plans without running the editor', async () => {
	const service = vi.fn(edit);
	const a = await setup(service),
		b = await setup();
	const token = a.captureTextReplaceToken('all-pages');
	const plan = a.planTextReplacement(token, { query: 'aa', replacement: 'X', mode: 'all' });
	expect(() => a.getTextReplaceOccurrences({} as ViewerTextReplaceToken, 'aa')).toThrow(
		/cancelled/,
	);
	await expect(b.applyTextReplacePlan(plan, token)).rejects.toMatchObject({ name: 'AbortError' });
	await expect(a.applyTextReplacePlan({ ...plan }, token)).rejects.toThrow(/owned search context/);
	expect(service).not.toHaveBeenCalled();
	a.destroy();
	b.destroy();
});
it('refuses reentrant host input getters before creating a usable plan', async () => {
	const c = await setup();
	const token = c.captureTextReplaceToken('current-page');
	expect(() =>
		c.planTextReplacement(token, {
			get query() {
				c.clearSelection();
				return 'aa';
			},
			replacement: 'X',
			mode: 'all',
		}),
	).toThrow(/cancelled/);
	c.destroy();
});
for (const event of ['selection-change', 'shape-select', 'page-change'] as const) {
	it(`lets ${event} callbacks override owned navigation without promoting new intent`, async () => {
		const c = await setup();
		const token = c.captureTextReplaceToken('all-pages');
		const match = c.getTextReplaceOccurrences(token, 'aa')[3]!;
		c.onEvent((name) => {
			if (name === event) c.clearSelection();
		});
		expect(() => c.selectTextReplaceOccurrence(token, match, 'aa')).toThrow(/cancelled/);
		expect(c.state.selectedShape).toBeNull();
		expect(c.isTextReplaceTokenCurrent(token)).toBe(false);
		c.destroy();
	});
}
it('refuses stale success after accepted document callbacks replace source', async () => {
	const c = await setup();
	const token = c.captureTextReplaceToken('all-pages');
	const plan = c.planTextReplacement(token, { query: 'aa', replacement: 'X', mode: 'all' });
	c.onEvent((name) => {
		if (name === 'document-change') c.setDocument(null);
	});
	await expect(c.applyTextReplacePlan(plan, token)).rejects.toMatchObject({ name: 'AbortError' });
	expect(c.state.document).toBeNull();
	c.destroy();
});
it('keeps current-page scope fixed while navigation changes the selected target', async () => {
	const c = await setup();
	let token = c.captureTextReplaceToken('current-page');
	const matches = c.getTextReplaceOccurrences(token, 'aa');
	token = c.selectTextReplaceOccurrence(token, matches[2]!, 'aa');
	expect(c.getTextReplaceOccurrences(token, 'aa')).toHaveLength(3);
	const plan = c.planTextReplacement(token, { query: 'aa', replacement: '', mode: 'all' });
	const next = await c.applyTextReplacePlan(plan, token);
	expect(c.state.document!.pages[1]!.shapes[0]!.text.plainText).toBe('😀aa');
	expect(c.getTextReplaceOccurrences(next, 'aa')).toEqual([]);
	c.destroy();
});
it('does not emit stale post-edit page events after the primary callback changes page', async () => {
	const c = await setup();
	const token = c.captureTextReplaceToken('all-pages');
	const matches = c.getTextReplaceOccurrences(token, 'aa');
	const plan = c.planTextReplacement(token, {
		query: 'aa',
		replacement: 'X',
		mode: 'current',
		current: matches[2]!,
	});
	const pages: number[] = [];
	c.onEvent((name, detail) => {
		if (name === 'shape-select') c.setPage(0);
		if (name === 'page-change') pages.push(detail as number);
	});
	await expect(c.applyTextReplacePlan(plan, token)).rejects.toMatchObject({ name: 'AbortError' });
	expect(pages).toEqual([0]);
	expect(c.state.pageIndex).toBe(0);
	expect(c.state.document!.pages[0]!.shapes[1]!.text.plainText).toBe('X');
	c.destroy();
});
it('refuses stale no-op results before navigating or publishing source', async () => {
	const gate = deferred();
	const c = await setup(async (bytes, commands) => {
		await gate.waiting;
		return edit(bytes, commands);
	});
	const before = c.exportVsdx().bytes;
	const token = c.captureTextReplaceToken('all-pages');
	const matches = c.getTextReplaceOccurrences(token, 'aa');
	const plan = c.planTextReplacement(token, {
		query: 'aa',
		replacement: 'aa',
		mode: 'current',
		current: matches[2]!,
	});
	const result = c.applyTextReplacePlan(plan, token);
	const failed = expect(result).rejects.toMatchObject({ name: 'AbortError' });
	c.clearSelection();
	gate.release();
	await failed;
	expect(c.state.pageIndex).toBe(0);
	expect(c.state.selectedShape).toBeNull();
	expect(c.exportVsdx().bytes).toEqual(before);
	c.destroy();
});
it('keeps unrelated search, zoom and clipboard readiness publications from invalidating scope', async () => {
	const c = await setup();
	const token = c.captureTextReplaceToken('selection');
	c.setSearchQuery('Unrelated');
	c.setZoom(2);
	await c.prepareClipboardSelection(c.captureClipboardToken());
	expect(c.isTextReplaceTokenCurrent(token)).toBe(true);
	expect(c.getTextReplaceOccurrences(token, 'aa')).toHaveLength(3);
	c.destroy();
});
