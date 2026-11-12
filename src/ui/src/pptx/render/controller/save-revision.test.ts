import { describe, expect, it, vi } from 'vitest';

import { createSaveRevisionController } from './save-revision';

function setup(initial: readonly unknown[] = [{}]) {
	let sources: readonly unknown[] = initial;
	const markClean = vi.fn();
	const discardRecovery = vi.fn();
	const controller = createSaveRevisionController({
		getSources: () => sources,
		serialize: async () => new Uint8Array([1, 2, 3]),
		markClean,
		discardRecovery,
	});
	return {
		controller,
		markClean,
		discardRecovery,
		edit: () => {
			sources = [{}];
		},
		setSources: (next: readonly unknown[]) => {
			sources = next;
		},
	};
}

describe('createSaveRevisionController', () => {
	it('keeps the revision until the sources change', () => {
		const { controller, edit } = setup();
		const first = controller.getRevision();
		expect(first).toBe(1);
		expect(controller.getRevision()).toBe(first);
		edit();
		expect(controller.getRevision()).toBe(first + 1);
	});

	it('tags a snapshot with the revision read before serialising', async () => {
		const { controller } = setup();
		const snapshot = await controller.snapshot();
		expect(snapshot.revision).toBe(1);
		expect([...snapshot.content]).toEqual([1, 2, 3]);
	});

	it('marks the document clean and drops recovery for the current revision', async () => {
		const { controller, markClean, discardRecovery } = setup();
		const { revision } = await controller.snapshot();
		expect(controller.acknowledge(revision)).toEqual({ status: 'clean', revision });
		expect(markClean).toHaveBeenCalledTimes(1);
		expect(discardRecovery).toHaveBeenCalledTimes(1);
		expect(controller.getSavedRevision()).toBe(revision);
	});

	it('keeps edits made while the save was in flight dirty and keeps recovery', async () => {
		const { controller, markClean, discardRecovery, edit } = setup();
		const { revision } = await controller.snapshot();
		edit();
		const result = controller.acknowledge(revision);
		expect(result).toEqual({ status: 'pending', revision, currentRevision: revision + 1 });
		expect(markClean).not.toHaveBeenCalled();
		expect(discardRecovery).not.toHaveBeenCalled();
		expect(controller.getSavedRevision()).toBe(revision);
	});

	it('ignores an acknowledgement older than one already accepted', async () => {
		const { controller, markClean, edit } = setup();
		const older = (await controller.snapshot()).revision;
		edit();
		const newer = (await controller.snapshot()).revision;
		expect(controller.acknowledge(newer).status).toBe('clean');
		markClean.mockClear();
		expect(controller.acknowledge(older)).toEqual({
			status: 'stale',
			revision: older,
			savedRevision: newer,
		});
		expect(markClean).not.toHaveBeenCalled();
		expect(controller.getSavedRevision()).toBe(newer);
	});

	it('does not repeat the recovery cleanup for a repeated acknowledgement', async () => {
		const { controller, markClean, discardRecovery } = setup();
		const { revision } = await controller.snapshot();
		controller.acknowledge(revision);
		controller.acknowledge(revision);
		expect(markClean).toHaveBeenCalledTimes(2);
		expect(discardRecovery).toHaveBeenCalledTimes(1);
	});

	it('rejects revisions that were never issued', () => {
		const { controller, markClean } = setup();
		const current = controller.getRevision();
		for (const bad of [0, -1, 1.5, Number.NaN, current + 1]) {
			expect(controller.acknowledge(bad).status).toBe('unknown');
		}
		expect(markClean).not.toHaveBeenCalled();
	});

	it('treats an edit then an undo to the same state as the same revision', async () => {
		const slides = {};
		let sources: readonly unknown[] = [slides];
		const controller = createSaveRevisionController({
			getSources: () => sources,
			serialize: async () => new Uint8Array(),
			markClean: () => undefined,
			discardRecovery: () => undefined,
		});
		const { revision } = await controller.snapshot();
		sources = [{}];
		sources = [slides];
		expect(controller.acknowledge(revision).status).toBe('clean');
	});

	it('never reports clean for a binding that supplies no change sources', async () => {
		const { controller, setSources, markClean } = setup();
		setSources([]);
		const { revision } = await controller.snapshot();
		expect(controller.acknowledge(revision).status).toBe('pending');
		expect(markClean).not.toHaveBeenCalled();
	});
});
