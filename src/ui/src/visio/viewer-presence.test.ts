import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import type { CollaborationRole } from 'ooxml-core/collab';
import { createVsdx, editVsdx, parseVsdx, type VisioEdit } from 'ooxml-core/visio';
import { sanitizeVisioPresence, VISIO_INITIAL_PRESENCE } from 'ooxml-core/visio/ui';
import { createBackstage } from './backstage';
import { ViewerController } from './controller';
import { registerViewerControls } from './office-ui';
import { renderPage } from './render-svg';
import { ViewerShare } from './viewer-share';
import type { VisioCollaborationOptions } from './collaboration-types';
import type { CancellableEditor } from './worker-editor';

const rect = (id: string, x: number): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId: id,
	x,
	y: 2,
	width: 2.8,
	height: 1,
});
const move = (id: string, x: number): VisioEdit => ({
	type: 'move-shape',
	pageId: '0',
	shapeId: id,
	x,
	y: 2,
});
const until = async (check: () => boolean) => {
	await vi.waitFor(() => expect(check()).toBe(true));
};

const disposers: (() => void)[] = [];
beforeEach(() => {
	// Ten client pixels are one drawing inch.
	Object.defineProperty(SVGSVGElement.prototype, 'getScreenCTM', {
		value: () => ({ a: 10, b: 0, inverse: () => ({}) }),
		configurable: true,
	});
	vi.stubGlobal(
		'DOMPoint',
		class {
			constructor(
				public x: number,
				public y: number,
			) {}
			matrixTransform() {
				return { x: this.x / 10, y: this.y / 10 };
			}
		},
	);
});
afterEach(() => {
	disposers.splice(0).forEach((dispose) => dispose());
	document.body.replaceChildren();
	delete (SVGSVGElement.prototype as { getScreenCTM?: unknown }).getScreenCTM;
	vi.unstubAllGlobals();
	vi.useRealTimers();
});

const hostSession = (
	hub: ReturnType<typeof createMemoryHub>,
	name: string,
	role?: CollaborationRole,
	sanitize = true,
) =>
	createCollabSession({
		roomId: 'presence',
		provider: transportProvider({ transport: hub.createTransport('presence') }),
		user: { name, color: name === 'Ada' ? '#16a34a' : '#dc2626', ...(role ? { role } : {}) },
		heartbeatMs: 0,
		syncGraceMs: 0,
		...(sanitize
			? { sanitizePayload: sanitizeVisioPresence, initialPresence: VISIO_INITIAL_PRESENCE }
			: {}),
	});

/** A window on real packages; `seed` opens a drawing with shapes 1 and 2, else it adopts the room's. */
async function viewerWindow(name: string, seed = true) {
	registerViewerControls();
	const editor: CancellableEditor = async (bytes, commands) => {
		const result = await editVsdx(bytes, commands);
		return { ...result, document: await parseVsdx(result.bytes) };
	};
	const controller = new ViewerController(parseVsdx, () => {}, editor);
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.append(createBackstage(document));
	const viewport = document.createElement('div');
	viewport.className = 'viewport';
	root.append(viewport);
	// The canvas: the page SVG is replaced whenever the document or page changes.
	let drawn: unknown;
	controller.subscribe((state) => {
		const page = state.document?.pages[state.pageIndex];
		if (!state.document || !page || drawn === state.document) return;
		drawn = state.document;
		viewport.replaceChildren(renderPage(state.document, page).svg);
	});
	const share = new ViewerShare(root, controller, () => ({
		displayName: name,
		avatarColor: '#16a34a',
	}));
	const dispose = share.wire();
	disposers.push(() => {
		dispose();
		controller.destroy();
	});
	if (seed) {
		await controller.load(await createVsdx());
		await controller.applyEdits([rect('1', 1), rect('2', 3)]);
	}
	return {
		controller,
		share,
		viewport,
		join: async (options: VisioCollaborationOptions) => {
			share.want(options);
			await vi.waitFor(() => expect(share.active).toBe(true));
		},
		move: (x: number, y: number) => {
			const event = new Event('pointermove', { bubbles: true });
			Object.defineProperties(event, { clientX: { value: x }, clientY: { value: y } });
			viewport.querySelector('svg')!.dispatchEvent(event);
		},
		cursors: () =>
			[...viewport.querySelectorAll<SVGGElement>('[data-presence-cursor]')].map((cursor) => ({
				name: cursor.querySelector('[data-presence-name]')!.textContent,
				fill: cursor.querySelector('path')!.getAttribute('fill'),
				transform: cursor.getAttribute('transform'),
			})),
		outlines: () =>
			[...viewport.querySelectorAll<SVGGElement>('[data-presence-selection]')].map((outline) => ({
				id: outline.dataset.presenceSelection,
				stroke: outline.querySelector('rect')!.getAttribute('stroke'),
			})),
	};
}
type Window = Awaited<ReturnType<typeof viewerWindow>>;

describe('live cursors and selections', () => {
	it('shows the pointer and the selection of a peer on the same page', async () => {
		const hub = createMemoryHub();
		const ada = await viewerWindow('Ada');
		const grace = await viewerWindow('Grace', false);
		await ada.join({ session: hostSession(hub, 'Ada') });
		await grace.join({ session: hostSession(hub, 'Grace') });
		await until(() => (grace.controller.state.document?.pages[0]?.shapes.length ?? 0) === 2);

		ada.move(30, 20);
		await vi.waitFor(() =>
			expect(grace.cursors()).toEqual([
				{ name: 'Ada', fill: '#16a34a', transform: 'translate(3 2) scale(0.1)' },
			]),
		);
		// Nobody sees their own cursor.
		expect(ada.cursors()).toEqual([]);

		ada.controller.selectShapes([{ id: '2', name: '', pageId: '0' }]);
		await vi.waitFor(() => expect(grace.outlines()).toEqual([{ id: '2', stroke: '#16a34a' }]));
		const outline = grace.viewport.querySelector('[data-presence-selection] rect')!;
		expect(outline.getAttribute('width')).toBe('2.8');

		// A pointer outside the page hides the cursor; clearing the selection removes the outline.
		ada.move(5000, 5000);
		ada.controller.clearSelection();
		await vi.waitFor(() => expect(grace.cursors()).toEqual([]));
		await vi.waitFor(() => expect(grace.outlines()).toEqual([]));
	});

	it('removes a peer that leaves', async () => {
		const hub = createMemoryHub();
		const ada = await viewerWindow('Ada');
		const grace = await viewerWindow('Grace', false);
		await ada.join({ session: hostSession(hub, 'Ada') });
		await grace.join({ session: hostSession(hub, 'Grace') });
		await until(() => (grace.controller.state.document?.pages[0]?.shapes.length ?? 0) === 2);
		ada.move(30, 20);
		await vi.waitFor(() => expect(grace.cursors()).toHaveLength(1));
		ada.share.stop();
		await vi.waitFor(() => expect(grace.cursors()).toEqual([]));
		expect(ada.viewport.querySelector('[data-presence-overlay]')).toBeNull();
	});

	it('does not crash when the host session was made without the Visio sanitiser', async () => {
		const hub = createMemoryHub();
		const ada = await viewerWindow('Ada');
		const grace = await viewerWindow('Grace', false);
		await ada.join({ session: hostSession(hub, 'Ada', undefined, false) });
		await grace.join({ session: hostSession(hub, 'Grace', undefined, false) });
		ada.move(30, 20);
		ada.controller.selectShapes([{ id: '1', name: '', pageId: '0' }]);
		await new Promise((resolve) => setTimeout(resolve, 120));
		expect(grace.cursors()).toEqual([]);
		expect(grace.outlines()).toEqual([]);
	});
});

describe('owner write-back', () => {
	let step = 10;
	const change = (w: Window, id = '1') => w.controller.applyEdits([move(id, (step += 1))]);
	const adopted = (w: Window) =>
		until(() => (w.controller.state.document?.pages[0]?.shapes.length ?? 0) === 2);

	it('fires once after the debounce for the owner, from local and remote changes', async () => {
		const hub = createMemoryHub();
		const owner = await viewerWindow('Ada');
		const other = await viewerWindow('Grace', false);
		const onWriteBack = vi.fn();
		await owner.join({
			session: hostSession(hub, 'Ada', 'owner'),
			onWriteBack,
			writeBackDebounceMs: 300,
		});
		await other.join({ session: hostSession(hub, 'Grace', 'collaborator') });
		await adopted(other);
		// Joining and adopting schedule saves of their own; let them settle before counting.
		await new Promise((resolve) => setTimeout(resolve, 800));
		onWriteBack.mockClear();
		await change(owner);
		await change(owner);
		expect(onWriteBack).not.toHaveBeenCalled();
		await vi.waitFor(() => expect(onWriteBack).toHaveBeenCalledTimes(1), { timeout: 5_000 });
		expect(onWriteBack.mock.calls[0]![0]).toBeInstanceOf(Uint8Array);
		await new Promise((resolve) => setTimeout(resolve, 800));
		expect(onWriteBack).toHaveBeenCalledTimes(1);

		// A collaborator's change reaches the owner as a remote step, which is saved too.
		await change(other, '2');
		await vi.waitFor(() => expect(onWriteBack).toHaveBeenCalledTimes(2), { timeout: 5_000 });
	});

	it('never fires for a collaborator', async () => {
		vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		const hub = createMemoryHub();
		const w = await viewerWindow('Grace');
		const onWriteBack = vi.fn();
		await w.join({
			session: hostSession(hub, 'Grace', 'collaborator'),
			onWriteBack,
			writeBackDebounceMs: 10,
		});
		await change(w);
		await vi.advanceTimersByTimeAsync(60_000);
		w.share.stop();
		expect(onWriteBack).not.toHaveBeenCalled();
	});

	it('uses a 5000 ms default and flushes a pending save on leave', async () => {
		vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		const hub = createMemoryHub();
		const w = await viewerWindow('Ada');
		const onWriteBack = vi.fn();
		await w.join({ session: hostSession(hub, 'Ada', 'owner'), onWriteBack });
		await change(w);
		await vi.advanceTimersByTimeAsync(4_900);
		expect(onWriteBack).not.toHaveBeenCalled();
		await vi.advanceTimersByTimeAsync(200);
		expect(onWriteBack).toHaveBeenCalledTimes(1);

		await change(w);
		expect(onWriteBack).toHaveBeenCalledTimes(1);
		w.share.stop();
		expect(onWriteBack).toHaveBeenCalledTimes(2);
		await vi.advanceTimersByTimeAsync(60_000);
		expect(onWriteBack).toHaveBeenCalledTimes(2);
	});
});

describe('owner write-back acknowledgement', () => {
	let step = 500;
	const change = (w: Window) => w.controller.applyEdits([move('1', (step += 1))]);

	it('reports saving then saved in the collaboration state once the promise resolves', async () => {
		vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		const hub = createMemoryHub();
		const w = await viewerWindow('Ada');
		let release!: () => void;
		const onWriteBack = vi.fn(
			() =>
				new Promise<void>((resolve) => {
					release = resolve;
				}),
		);
		await w.join({
			session: hostSession(hub, 'Ada', 'owner'),
			onWriteBack,
			writeBackDebounceMs: 20,
		});
		await change(w);
		expect(w.share.state.writeBack?.status).toBe('pending');
		await vi.advanceTimersByTimeAsync(30);
		expect(onWriteBack).toHaveBeenCalledTimes(1);
		expect(w.share.state.writeBack?.status).toBe('saving');
		release();
		await vi.advanceTimersByTimeAsync(0);
		expect(w.share.state.writeBack?.status).toBe('saved');
		expect(w.share.state.writeBack?.lastSavedAt).toBeTypeOf('number');
		w.share.stop();
	});

	it('retries a rejection, then reports saved', async () => {
		vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		const hub = createMemoryHub();
		const w = await viewerWindow('Ada');
		const onWriteBack = vi
			.fn<(bytes: Uint8Array) => Promise<void>>()
			.mockRejectedValueOnce(new Error('offline'))
			.mockResolvedValue(undefined);
		await w.join({
			session: hostSession(hub, 'Ada', 'owner'),
			onWriteBack,
			writeBackDebounceMs: 20,
			writeBackRetry: { attempts: 3, baseDelayMs: 100 },
		});
		await change(w);
		await vi.advanceTimersByTimeAsync(30);
		expect(onWriteBack).toHaveBeenCalledTimes(1);
		await vi.advanceTimersByTimeAsync(120);
		expect(onWriteBack).toHaveBeenCalledTimes(2);
		expect(w.share.state.writeBack?.status).toBe('saved');
		w.share.stop();
	});

	it('reports failed with the error when the retries are exhausted', async () => {
		vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		const hub = createMemoryHub();
		const w = await viewerWindow('Ada');
		const states: unknown[] = [];
		w.share.onChange = (state) => states.push(state.writeBack?.status);
		const onWriteBack = vi.fn(() => Promise.reject(new Error('quota')));
		await w.join({
			session: hostSession(hub, 'Ada', 'owner'),
			onWriteBack,
			writeBackDebounceMs: 20,
			writeBackRetry: { attempts: 2, baseDelayMs: 50 },
		});
		await change(w);
		await vi.advanceTimersByTimeAsync(500);
		expect(onWriteBack).toHaveBeenCalledTimes(2);
		expect(w.share.state.writeBack).toMatchObject({ status: 'failed', error: 'quota' });
		expect(states).toContain('failed');
		w.share.stop();
	});

	it('flushes a pending save on leave and never reports write-back for a collaborator', async () => {
		vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		const hub = createMemoryHub();
		const w = await viewerWindow('Ada');
		const onWriteBack = vi.fn(() => Promise.resolve());
		await w.join({ session: hostSession(hub, 'Ada', 'owner'), onWriteBack });
		await change(w);
		w.share.stop();
		expect(onWriteBack).toHaveBeenCalledTimes(1);
		await vi.advanceTimersByTimeAsync(60_000);
		expect(onWriteBack).toHaveBeenCalledTimes(1);

		const g = await viewerWindow('Grace');
		await g.join({ session: hostSession(hub, 'Grace', 'collaborator'), onWriteBack });
		expect(g.share.state.writeBack).toBeUndefined();
		g.share.stop();
	});
});
