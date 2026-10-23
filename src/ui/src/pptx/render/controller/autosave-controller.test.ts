import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { AutosaveControllerDeps } from './autosave-controller';
import { createAutosaveController } from './autosave-controller';

const INTERVAL = 1000;

function setup(overrides: Partial<AutosaveControllerDeps> = {}) {
	let active = true;
	let filePath: string | undefined = 'deck.pptx';
	let held = false;
	const writes: Uint8Array[] = [];
	const writeSnapshot = vi.fn(async (_key: string, bytes: Uint8Array) => {
		writes.push(bytes);
	});
	const serialize = vi.fn(async () => new Uint8Array([1, 2, 3]));
	const onSaved = vi.fn();
	const controller = createAutosaveController({
		getActivation: () =>
			active ? { active: true } : { active: false, reason: 'autosave_toggle_off' },
		getFilePath: () => filePath,
		getIntervalMs: () => INTERVAL,
		serialize,
		isHeld: () => held,
		onSaved,
		writeSnapshot,
		...overrides,
	});
	return {
		controller,
		writes,
		writeSnapshot,
		serialize,
		onSaved,
		setActive: (next: boolean) => {
			active = next;
			controller.refresh();
		},
		setFilePath: (next: string | undefined) => {
			filePath = next;
			controller.refresh();
		},
		setHeld: (next: boolean) => {
			held = next;
		},
	};
}

describe('createAutosaveController', () => {
	beforeEach(() => {
		vi.useFakeTimers();
	});
	afterEach(() => {
		vi.useRealTimers();
	});

	it('writes one interval after the first unsaved edit, however many edits follow', async () => {
		const { controller, writeSnapshot, onSaved } = setup();
		controller.markEdited();
		await vi.advanceTimersByTimeAsync(400);
		controller.markEdited();
		await vi.advanceTimersByTimeAsync(400);
		controller.markEdited();
		expect(writeSnapshot).not.toHaveBeenCalled();
		await vi.advanceTimersByTimeAsync(200);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
		expect(onSaved).toHaveBeenCalledTimes(1);
		expect(controller.getSnapshot()).toMatchObject({
			dirty: false,
			status: { state: 'saved' },
		});
	});

	it('writes nothing while the document is unchanged', async () => {
		const { controller, writeSnapshot } = setup();
		await vi.advanceTimersByTimeAsync(INTERVAL * 5);
		expect(writeSnapshot).not.toHaveBeenCalled();
		controller.markEdited();
		await vi.advanceTimersByTimeAsync(INTERVAL * 5);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
	});

	it('reports disabled with the activation reason, and resumes on refresh', async () => {
		const { controller, setActive, writeSnapshot } = setup();
		setActive(false);
		expect(controller.getSnapshot().status).toEqual({
			state: 'disabled',
			reason: 'autosave_toggle_off',
		});
		controller.markEdited();
		await vi.advanceTimersByTimeAsync(INTERVAL * 2);
		expect(writeSnapshot).not.toHaveBeenCalled();
		setActive(true);
		expect(controller.getSnapshot().status).toEqual({ state: 'idle' });
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
	});

	it('is disabled without a file key', () => {
		const { controller, setFilePath } = setup({
			getActivation: () => ({ active: true }),
		});
		setFilePath(undefined);
		expect(controller.getSnapshot().status).toEqual({ state: 'disabled', reason: 'no_file_path' });
	});

	it('saveNow writes even when clean, and markClean cancels a pending write', async () => {
		const { controller, writeSnapshot } = setup();
		await controller.saveNow();
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
		controller.markEdited();
		controller.markClean();
		await vi.advanceTimersByTimeAsync(INTERVAL * 2);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
	});

	it('keeps an edit made during a write for the next interval', async () => {
		let release: (() => void) | undefined;
		const { controller, writeSnapshot } = setup({
			serialize: () =>
				new Promise((resolve) => {
					release = () => resolve(new Uint8Array([9]));
				}),
		});
		controller.markEdited();
		await vi.advanceTimersByTimeAsync(INTERVAL);
		controller.markEdited();
		release?.();
		await vi.advanceTimersByTimeAsync(0);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
		expect(controller.getSnapshot().dirty).toBe(true);
		release = undefined;
		await vi.advanceTimersByTimeAsync(INTERVAL);
		release?.();
		await vi.advanceTimersByTimeAsync(0);
		expect(writeSnapshot).toHaveBeenCalledTimes(2);
	});

	it('defers a due write while a preview is held', async () => {
		const { controller, writeSnapshot, setHeld } = setup();
		controller.markEdited();
		setHeld(true);
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(writeSnapshot).not.toHaveBeenCalled();
		setHeld(false);
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
	});

	it('reports a failed write as an error and keeps the edits outstanding', async () => {
		const { controller } = setup({
			writeSnapshot: async () => {
				throw new Error('quota');
			},
		});
		controller.markEdited();
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(controller.getSnapshot()).toMatchObject({
			dirty: true,
			status: { state: 'error', message: 'quota' },
		});
	});

	it('observe: clean clears, a changed source is an edit, an unchanged one is not', async () => {
		const { controller, writeSnapshot } = setup();
		const slides = [{}];
		controller.observe(false, [slides]);
		controller.observe(true, [slides]);
		expect(controller.getSnapshot().dirty).toBe(true);
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
		// Still dirty in the editor (no real save), nothing reassigned: no rewrite.
		controller.observe(true, [slides]);
		await vi.advanceTimersByTimeAsync(INTERVAL * 3);
		expect(writeSnapshot).toHaveBeenCalledTimes(1);
		controller.observe(true, [[...slides]]);
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(writeSnapshot).toHaveBeenCalledTimes(2);
		controller.observe(true, [[]]);
		controller.observe(false, [[]]);
		await vi.advanceTimersByTimeAsync(INTERVAL * 2);
		expect(writeSnapshot).toHaveBeenCalledTimes(2);
	});

	it('markClean with saved shows the real save', () => {
		const { controller } = setup();
		controller.markEdited();
		controller.markClean({ saved: true });
		expect(controller.getSnapshot()).toMatchObject({ dirty: false, status: { state: 'saved' } });
	});

	it('writes nothing after destroy', async () => {
		const { controller, writeSnapshot } = setup();
		controller.markEdited();
		controller.destroy();
		await vi.advanceTimersByTimeAsync(INTERVAL * 2);
		expect(writeSnapshot).not.toHaveBeenCalled();
	});
});

describe('createAutosaveController inline drafts', () => {
	beforeEach(() => {
		vi.useFakeTimers();
	});
	afterEach(() => {
		vi.useRealTimers();
	});

	it('counts typed-but-uncommitted text as an edit and hands it to the snapshot', async () => {
		let editing: string | null = 'title';
		const serialize = vi.fn(async (_draft: unknown) => new Uint8Array([1]));
		const controller = createAutosaveController({
			getActivation: () => ({ active: true }),
			getFilePath: () => 'deck.pptx',
			getIntervalMs: () => INTERVAL,
			serialize,
			getInlineEditingId: () => editing,
			writeSnapshot: async () => undefined,
		});
		controller.noteInlineDraft({ slideId: 's1', elementId: 'title', text: 'Hel' });
		controller.noteInlineDraft({ slideId: 's1', elementId: 'title', text: 'Hello' });
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(serialize).toHaveBeenCalledExactlyOnceWith({
			slideId: 's1',
			elementId: 'title',
			text: 'Hello',
		});
		// The same draft again is not a new edit.
		controller.noteInlineDraft({ slideId: 's1', elementId: 'title', text: 'Hello' });
		await vi.advanceTimersByTimeAsync(INTERVAL * 2);
		expect(serialize).toHaveBeenCalledOnce();
		// A cancelled edit never reaches a snapshot.
		controller.noteInlineDraft({ slideId: 's1', elementId: 'title', text: 'Hello!' });
		editing = null;
		await vi.advanceTimersByTimeAsync(INTERVAL);
		expect(serialize).toHaveBeenLastCalledWith(null);
	});
});
