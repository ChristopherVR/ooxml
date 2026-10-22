// Remote changes reach the open drawing by reading only the pages that changed, and by reading the
// whole package otherwise. The parser is a spy, so the test counts what was really read.
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import { createVsdx, editVsdx, parseVsdx, type VisioEdit } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { DrawingSync, type DrawingCollab } from './viewer-share-sync';
import type { CancellableEditor } from './worker-editor';

const cleanup: (() => void)[] = [];
afterEach(() => cleanup.splice(0).forEach((stop) => stop()));

const until = async (check: () => boolean | Promise<boolean>) => {
	for (let i = 0; i < 300 && !(await check()); i++) await new Promise((r) => setTimeout(r, 10));
	expect(await check()).toBe(true);
};
const rect = (id: string, x: number, pageId = '0'): VisioEdit => ({
	type: 'create-rectangle',
	pageId,
	shapeId: id,
	x,
	y: 2,
	width: 1,
	height: 1,
});
const move = (id: string, x: number, pageId = '0'): VisioEdit => ({
	type: 'move-shape',
	pageId,
	shapeId: id,
	x,
	y: 2,
});

interface Peer {
	controller: ViewerController;
	sync: DrawingSync;
	full: ReturnType<typeof vi.fn>;
	partial: ReturnType<typeof vi.fn>;
	errors: string[];
}
const pin = (peer: Peer, pageId: string, id: string) =>
	peer.controller.state.document?.pages
		.find((page) => page.id === pageId)
		?.shapes.find((shape) => shape.id === id)?.transform[4];

async function startDrawing(): Promise<Uint8Array> {
	let bytes = (await editVsdx(await createVsdx(), [rect('1', 1), rect('2', 3)])).bytes;
	bytes = (
		await editVsdx(bytes, [{ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'Two' }])
	).bytes;
	return (await editVsdx(bytes, [rect('1', 2, '5')])).bytes;
}

/** A parser that counts whole reads and page reads. */
function spyParser() {
	const full = vi.fn((bytes: Uint8Array | ArrayBuffer) => parseVsdx(bytes));
	const partial = vi.fn((bytes: Uint8Array | ArrayBuffer, pages: readonly string[]) =>
		parseVsdx(bytes, { pages }),
	);
	return {
		full,
		partial,
		parser: Object.assign((bytes: Uint8Array | ArrayBuffer) => full(bytes), { pages: partial }),
	};
}

const editor: CancellableEditor = async (bytes, commands) => {
	const result = await editVsdx(bytes, commands);
	return { ...result, document: await parseVsdx(result.bytes) };
};

async function peer(
	drawing: DrawingCollab,
	hub: ReturnType<typeof createMemoryHub>,
	initial: Uint8Array,
	join: boolean,
): Promise<Peer> {
	const { full, partial, parser } = spyParser();
	const controller = new ViewerController(parser, () => {}, editor);
	await controller.load(initial);
	const session: CollabSession<object> = createCollabSession({
		roomId: 'drawing',
		provider: transportProvider({ transport: hub.createTransport('drawing') }),
		user: { name: join ? 'Bob' : 'Ada' },
		heartbeatMs: 0,
		syncGraceMs: 0,
	});
	const errors: string[] = [];
	const sync = new DrawingSync(drawing, session, controller, {
		join,
		fail: (message) => errors.push(message),
	});
	cleanup.push(() => {
		sync.dispose();
		session.destroy();
		controller.destroy();
	});
	return { controller, sync, full, partial, errors };
}

async function pair() {
	const drawing = await import('ooxml-core/visio/collab');
	const hub = createMemoryHub();
	const a = await peer(drawing, hub, await startDrawing(), false);
	await a.sync.idle();
	const b = await peer(drawing, hub, await createVsdx(), true);
	await until(() => b.controller.state.document?.pages.length === 2);
	await a.sync.idle();
	await b.sync.idle();
	for (const p of [a, b]) {
		p.full.mockClear();
		p.partial.mockClear();
	}
	return { drawing, hub, a, b };
}

describe('applying a shared change to the open drawing', () => {
	it('reads only the page a collaborator edited and splices it into the open document', async () => {
		const { a, b } = await pair();
		const untouched = b.controller.state.document!.pages.find((page) => page.id === '0');
		await a.controller.applyEdits([move('1', 6, '5')]);
		await until(() => pin(b, '5', '1') === 5.5);
		await b.sync.idle();
		expect(b.partial).toHaveBeenCalledTimes(1);
		expect(b.partial.mock.calls[0]![1]).toEqual(['5']);
		expect(b.full).not.toHaveBeenCalled();
		// The page that did not change is the object already on screen.
		expect(b.controller.state.document!.pages.find((page) => page.id === '0')).toBe(untouched);
		expect(b.errors).toEqual([]);
	});

	it('leaves the document and the saved bytes as a whole read of that package would', async () => {
		const { drawing, hub, a, b } = await pair();
		await a.controller.applyEdits([move('1', 6, '5'), rect('7', 5, '0')]);
		await until(() => pin(b, '5', '1') === 5.5 && pin(b, '0', '7') !== undefined);
		await b.sync.idle();
		expect(b.partial).toHaveBeenCalled();
		const saved = b.controller.exportVsdx().bytes;
		expect(b.controller.state.document).toEqual(await parseVsdx(saved));
		// A window that adopts the room from scratch reads it whole and ends up with the same model.
		const c = await peer(drawing, hub, await createVsdx(), true);
		await until(() => pin(c, '0', '7') !== undefined);
		await c.sync.idle();
		expect(c.partial).not.toHaveBeenCalled();
		expect(await parseVsdx(c.controller.exportVsdx().bytes)).toEqual(
			await parseVsdx(b.controller.exportVsdx().bytes),
		);
	});

	it('reads the whole package when the page list changes', async () => {
		const { a, b } = await pair();
		await a.controller.applyEdits([{ type: 'rename-page', pageId: '5', name: 'Renamed' }]);
		await until(() => b.controller.state.document?.pages[1]?.name === 'Renamed');
		await b.sync.idle();
		expect(b.partial).not.toHaveBeenCalled();
		expect(b.full).toHaveBeenCalledTimes(1);
	});

	it('reads the whole package after an edit made in this window', async () => {
		const { a, b } = await pair();
		await b.controller.applyEdits([move('2', 4)]);
		await until(() => pin(a, '0', '2') === 3.5);
		await a.sync.idle();
		await b.sync.idle();
		b.full.mockClear();
		b.partial.mockClear();
		await a.controller.applyEdits([move('1', 6, '5')]);
		await until(() => pin(b, '5', '1') === 5.5);
		await b.sync.idle();
		expect(b.partial).not.toHaveBeenCalled();
		expect(b.full).toHaveBeenCalledTimes(1);
	});
});

describe('the controller falls back to a whole read', () => {
	async function loaded() {
		const bytes = await startDrawing();
		const { full, partial, parser } = spyParser();
		const controller = new ViewerController(parser);
		await controller.load(bytes);
		cleanup.push(() => controller.destroy());
		const edited = (await editVsdx(bytes, [move('1', 6, '5')])).bytes;
		full.mockClear();
		return { controller, full, partial, edited };
	}

	it('uses the pages read when they fit', async () => {
		const { controller, full, partial, edited } = await loaded();
		await controller.applyRemoteSource(edited, { pages: ['5'] });
		expect(partial).toHaveBeenCalledTimes(1);
		expect(full).not.toHaveBeenCalled();
		expect(controller.state.document).toEqual(await parseVsdx(edited));
	});

	it('reads everything without a change, or from a parser that cannot read pages', async () => {
		const { controller, full, partial, edited } = await loaded();
		await controller.applyRemoteSource(edited);
		expect(partial).not.toHaveBeenCalled();
		expect(full).toHaveBeenCalledTimes(1);
		const plain = new ViewerController((input) => parseVsdx(input));
		cleanup.push(() => plain.destroy());
		await plain.load(await startDrawing());
		await plain.applyRemoteSource(edited, { pages: ['5'] });
		expect(plain.state.document).toEqual(await parseVsdx(edited));
	});

	it('reads everything when the pages do not merge or the page read fails', async () => {
		const { controller, full, partial, edited } = await loaded();
		await controller.applyRemoteSource(edited, { pages: ['missing'] });
		expect(partial).toHaveBeenCalledTimes(1);
		expect(full).toHaveBeenCalledTimes(1);
		partial.mockRejectedValueOnce(new Error('worker failed'));
		await controller.applyRemoteSource((await editVsdx(edited, [move('1', 7, '5')])).bytes, {
			pages: ['5'],
		});
		expect(full).toHaveBeenCalledTimes(2);
		expect(controller.state.error).toBeNull();
	});

	it('keeps a cancelled read cancelled', async () => {
		const { controller, full, partial, edited } = await loaded();
		partial.mockRejectedValueOnce(new DOMException('cancelled', 'AbortError'));
		await expect(controller.applyRemoteSource(edited, { pages: ['5'] })).rejects.toThrow(
			/cancelled/,
		);
		expect(full).not.toHaveBeenCalled();
	});
});
