import { expect, it, vi } from 'vitest';
import { createBuildingBlockSession } from '../src/building-blocks';
it('source read replacement ignores stale failures and publishes only the latest error', async () => {
	const session = createBuildingBlockSession();
	let reject!: (reason: Error) => void;
	const slow = new Blob();
	slow.arrayBuffer = () =>
		new Promise((_, fail) => {
			reject = fail;
		});
	const error = vi.fn();
	session.update({ content: slow, onDocumentError: error });
	expect(session.controller.state.loading).toBe(true);
	const latest = new Error('latest source');
	const next = new Blob();
	next.arrayBuffer = async () => {
		throw latest;
	};
	session.update({ content: next, onDocumentError: error });
	await vi.waitFor(() => expect(session.controller.state.error).toBe(latest));
	reject(new Error('stale source'));
	await Promise.resolve();
	expect(session.controller.state.error).toBe(latest);
	expect(error).toHaveBeenCalledExactlyOnceWith(latest);
	session.destroy();
	expect(() => session.controller.setZoom(2)).toThrow();
});
