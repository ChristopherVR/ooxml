import { describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import { YDOC_ASSETS_KEY } from '../collaboration-assets';
import { LOCAL_SYNC_ORIGIN, reconcileSlidesInYDoc } from '../collaboration-reconcile';
import { readSlidesFromYDoc, YDOC_SLIDES_KEY } from '../collaboration-sync';
import {
	bindCollaborationHistory,
	collaborationHistoryReady,
	createCollaborationHistory,
} from './collaboration-history';

describe('collaboration history', () => {
	it('distinguishes editors that share a host-owned document', () => {
		const doc = new Y.Doc();
		const shape = new Y.Map<string>();
		doc.getArray(YDOC_SLIDES_KEY).push([shape]);
		const alice = {};
		const bob = {};
		const history = createCollaborationHistory(doc, Y.UndoManager, alice);
		try {
			doc.transact(() => shape.set('alice', 'one'), alice);
			doc.transact(() => shape.set('bob', 'two'), bob);
			history.undo();
			expect(shape.get('alice')).toBeUndefined();
			expect(shape.get('bob')).toBe('two');
		} finally {
			history.destroy();
			doc.destroy();
		}
	});
	it('does not record the initial local room seed', () => {
		const doc = new Y.Doc();
		const history = createCollaborationHistory(doc, Y.UndoManager, LOCAL_SYNC_ORIGIN);
		try {
			const slide = new Y.Map<string>();
			doc.transact(() => doc.getArray(YDOC_SLIDES_KEY).push([slide]), LOCAL_SYNC_ORIGIN);
			expect(history.getSnapshot().canUndo).toBe(false);
			doc.transact(() => slide.set('name', 'Edit'), LOCAL_SYNC_ORIGIN);
			expect(history.getSnapshot().canUndo).toBe(true);
			history.undo();
			expect(slide.get('name')).toBeUndefined();
			expect(doc.getArray(YDOC_SLIDES_KEY).length).toBe(1);
		} finally {
			history.destroy();
			doc.destroy();
		}
	});
	it('preserves a peer shape change when undoing a reconciled slide edit', () => {
		const alice = new Y.Doc();
		const bob = new Y.Doc();
		const factories = {
			createArray: () => new Y.Array(),
			createMap: () => new Y.Map(),
			createText: () => new Y.Text(),
		};
		const slide = {
			id: 's1',
			slideNumber: 1,
			elements: [
				{ id: 'a', type: 'text' as const, x: 0, y: 0, width: 100, height: 50, text: 'Alpha' },
				{ id: 'b', type: 'text' as const, x: 0, y: 60, width: 100, height: 50, text: 'Beta' },
			],
		};
		try {
			reconcileSlidesInYDoc([slide], alice, factories);
			Y.applyUpdate(bob, Y.encodeStateAsUpdate(alice));
			const history = createCollaborationHistory(alice, Y.UndoManager, LOCAL_SYNC_ORIGIN);
			try {
				const local = readSlidesFromYDoc(alice);
				local[0].elements[0].x = 20;
				reconcileSlidesInYDoc(local, alice, factories);
				const remote = readSlidesFromYDoc(bob);
				remote[0].elements[1].y = 90;
				reconcileSlidesInYDoc(remote, bob, factories);
				Y.applyUpdate(alice, Y.encodeStateAsUpdate(bob), 'remote');
				history.undo();
				const result = readSlidesFromYDoc(alice)[0];
				expect(result.elements[0].x).toBe(0);
				expect(result.elements[1].y).toBe(90);
			} finally {
				history.destroy();
			}
		} finally {
			alice.destroy();
			bob.destroy();
		}
	});
	it('undoes local text without deleting a peer insertion in the same text', () => {
		const alice = new Y.Doc();
		const bob = new Y.Doc();
		try {
			const text = new Y.Text('Alpha');
			alice.getArray(YDOC_SLIDES_KEY).push([text]);
			Y.applyUpdate(bob, Y.encodeStateAsUpdate(alice));
			const history = createCollaborationHistory(alice, Y.UndoManager, LOCAL_SYNC_ORIGIN);
			try {
				alice.transact(() => text.insert(5, ' Alice'), LOCAL_SYNC_ORIGIN);
				const remote = bob.getArray<Y.Text>(YDOC_SLIDES_KEY).get(0);
				bob.transact(() => remote.insert(0, 'Bob '), LOCAL_SYNC_ORIGIN);
				Y.applyUpdate(alice, Y.encodeStateAsUpdate(bob), 'remote');
				history.undo();
				expect(text.toString()).toBe('Bob Alpha');
				history.redo();
				expect(text.toString()).toBe('Bob Alpha Alice');
			} finally {
				history.destroy();
			}
		} finally {
			alice.destroy();
			bob.destroy();
		}
	});

	it('undoes only local edits and preserves a peer edit on another shape', () => {
		const alice = new Y.Doc();
		const bob = new Y.Doc();
		try {
			const a = new Y.Map<string>();
			const b = new Y.Map<string>();
			a.set('text', 'Alpha');
			b.set('text', 'Beta');
			alice.getArray(YDOC_SLIDES_KEY).push([a, b]);
			Y.applyUpdate(bob, Y.encodeStateAsUpdate(alice));
			const history = createCollaborationHistory(alice, Y.UndoManager, LOCAL_SYNC_ORIGIN);
			try {
				alice.transact(() => a.set('text', 'Alpha Alice'), LOCAL_SYNC_ORIGIN);
				Y.applyUpdate(bob, Y.encodeStateAsUpdate(alice));
				const peerB = bob.getArray<Y.Map<string>>(YDOC_SLIDES_KEY).get(1);
				bob.transact(() => peerB.set('text', 'Beta Bob'), LOCAL_SYNC_ORIGIN);
				Y.applyUpdate(alice, Y.encodeStateAsUpdate(bob), 'remote');
				expect(history.getSnapshot().canUndo).toBe(true);
				history.undo();
				expect(a.get('text')).toBe('Alpha');
				expect(b.get('text')).toBe('Beta Bob');
				history.redo();
				expect(a.get('text')).toBe('Alpha Alice');
				expect(b.get('text')).toBe('Beta Bob');
			} finally {
				history.destroy();
			}
		} finally {
			alice.destroy();
			bob.destroy();
		}
	});

	it('keeps asset changes in the same undo step and respects explicit boundaries', () => {
		const doc = new Y.Doc();
		try {
			const shape = new Y.Map<string>();
			doc.getArray(YDOC_SLIDES_KEY).push([shape]);
			const assets = doc.getMap<string>(YDOC_ASSETS_KEY);
			const history = createCollaborationHistory(doc, Y.UndoManager, LOCAL_SYNC_ORIGIN);
			try {
				doc.transact(() => {
					shape.set('ref', 'image');
					assets.set('image', 'one');
				}, LOCAL_SYNC_ORIGIN);
				history.stopCapturing();
				doc.transact(() => assets.set('image', 'two'), LOCAL_SYNC_ORIGIN);
				history.undo();
				expect(assets.get('image')).toBe('one');
				history.undo();
				expect(shape.get('ref')).toBeUndefined();
				expect(assets.get('image')).toBeUndefined();
				expect(history.getSnapshot()).toEqual({ canUndo: false, canRedo: true });
			} finally {
				history.destroy();
			}
		} finally {
			doc.destroy();
		}
	});
});

describe('bindCollaborationHistory', () => {
	it('attaches a room history once yjs loads and detaches on release', async () => {
		const doc = new Y.Doc();
		const attach = vi.fn((history: unknown) => (history ? () => detach() : () => {}));
		const detach = vi.fn();
		const release = bindCollaborationHistory({ attachCollaboration: attach }, doc, async () => Y);
		await Promise.resolve();
		await Promise.resolve();
		expect(attach).toHaveBeenCalledTimes(1);
		release();
		expect(detach).toHaveBeenCalledTimes(1);
		doc.destroy();
	});

	it('does nothing when released before yjs resolves or without an UndoManager', async () => {
		const doc = new Y.Doc();
		const attach = vi.fn(() => () => {});
		bindCollaborationHistory({ attachCollaboration: attach }, doc, async () => ({}));
		const early = bindCollaborationHistory({ attachCollaboration: attach }, doc, async () => Y);
		early();
		await Promise.resolve();
		await Promise.resolve();
		expect(attach).not.toHaveBeenCalled();
		doc.destroy();
	});
});

describe('collaborationHistoryReady', () => {
	const base = { hasDoc: true, external: false, connected: true, synced: true, viewer: false };
	it('needs a writable, synchronized room unless the host owns the document', () => {
		expect(collaborationHistoryReady(base)).toBe(true);
		expect(collaborationHistoryReady({ ...base, synced: false })).toBe(false);
		expect(collaborationHistoryReady({ ...base, connected: false })).toBe(false);
		expect(collaborationHistoryReady({ ...base, viewer: true })).toBe(false);
		expect(collaborationHistoryReady({ ...base, hasDoc: false })).toBe(false);
		expect(
			collaborationHistoryReady({ ...base, external: true, connected: false, synced: false }),
		).toBe(true);
	});
});
