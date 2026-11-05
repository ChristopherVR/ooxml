import type { MasterViewCrudAction, MasterViewCrudActionId, MasterViewTarget } from 'ooxml-ui/pptx';
import type { PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createInitialViewerState, createStore } from '../state';
import type { Store, ViewerState } from '../state';
import { createMasterViewCrudActions } from './editor-master-view-crud-actions';
import { createEditorOps } from './editor-operations';
import type { EditorOps } from './editor-operations';

/** In-app dialogs load on demand; a busy CI runner needs more than waitFor's 1 s default. */
const DIALOG_WAIT = { timeout: 5000 };

const { applyMasterViewCrudAction, masterViewCrudActions } = vi.hoisted(() => ({
	applyMasterViewCrudAction: vi.fn(),
	masterViewCrudActions: vi.fn(),
}));

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => {
	const actual = await importOriginal();
	return { ...actual, applyMasterViewCrudAction, masterViewCrudActions };
});

/**
 * B4: the sidebar CRUD commands ZIP-surgery through
 * `pptx-viewer-shared/render/master-view-crud`, which is unit-tested on its
 * own; this exercises vanilla's WIRING around it, the part unique to this
 * binding: adopting the fresh handler + resolved master data with a history
 * entry, prompting for a new name, and surfacing a rejection.
 */

function harness(overrides: Partial<ViewerState> = {}): {
	store: Store<ViewerState>;
	setHandler: ReturnType<typeof vi.fn>;
	handler: PptxHandler;
	run: (id: MasterViewCrudActionId) => Promise<void>;
	ops: EditorOps;
} {
	const store = createStore<ViewerState>({
		...createInitialViewerState(),
		editable: true,
		slideMasters: [{ path: 'master-1', name: 'Corporate', elements: [], layouts: [] }],
		masterViewTab: 'slides',
		masterViewTarget: { masterIndex: 0, layoutIndex: null },
		...overrides,
	});
	const setHandler = vi.fn();
	const handler = { getImageData: vi.fn() } as unknown as PptxHandler;
	const ops = createEditorOps({ store, getHandler: () => handler, onHistoryChange: vi.fn() });
	const actions = createMasterViewCrudActions({
		doc: document,
		getTranslator: () => createTranslator(),
		store,
		ops,
		getHandler: () => handler,
		setHandler,
	});
	return { store, setHandler, handler, run: actions.runMasterViewCrudAction, ops };
}

/** The in-app dialog the code under test opened; waits for it to appear. */
async function openDialog(): Promise<ShadowRoot> {
	const dialog = await vi.waitFor(() => {
		const found = document.querySelector('office-ui-message-dialog');
		if (!found?.shadowRoot) throw new Error('no dialog yet');
		return found;
	}, DIALOG_WAIT);
	return dialog.shadowRoot as ShadowRoot;
}

/** Type into the open dialog's input (when given), then press a button. */
async function answerDialog(button: 'confirm' | 'cancel', value?: string): Promise<string> {
	const root = await openDialog();
	const message = root.querySelector('.message')?.textContent ?? '';
	const input = root.querySelector('input');
	if (input && value !== undefined) input.value = value;
	(root.querySelector(`button.${button}`) as HTMLButtonElement).click();
	await new Promise((resolve) => setTimeout(resolve, 0));
	return message;
}

function stubActionList(id: string, enabled = true): void {
	masterViewCrudActions.mockReturnValue([
		{ id, labelKey: `pptx.masterView.${id}`, enabled } as MasterViewCrudAction,
	]);
}

describe('vanilla master-view CRUD action wiring', () => {
	beforeEach(() => {
		applyMasterViewCrudAction.mockReset();
		masterViewCrudActions.mockReset();
	});
	afterEach(() => {
		document.querySelectorAll('office-ui-message-dialog').forEach((node) => node.remove());
		vi.restoreAllMocks();
	});

	it('does nothing outside master view', async () => {
		const { run } = harness({ masterViewTarget: null });
		stubActionList('addLayout');
		await run('addLayout');
		expect(applyMasterViewCrudAction).not.toHaveBeenCalled();
	});

	it('does nothing on a read-only deck', async () => {
		const { run } = harness({ editable: false });
		stubActionList('addLayout');
		await run('addLayout');
		expect(applyMasterViewCrudAction).not.toHaveBeenCalled();
	});

	it('does nothing when the action is disabled for the current target', async () => {
		const { run } = harness();
		stubActionList('deleteMaster', false);
		await run('deleteMaster');
		expect(applyMasterViewCrudAction).not.toHaveBeenCalled();
	});

	it('adopts the returned handler + slide masters with a history entry on success', async () => {
		const { run, store, setHandler, handler } = harness();
		stubActionList('addLayout');
		const newHandler = { getImageData: vi.fn() } as unknown as PptxHandler;
		const target: MasterViewTarget = { tab: 'slides', masterIndex: 0, layoutIndex: 1 };
		applyMasterViewCrudAction.mockResolvedValue({
			ok: true,
			handler: newHandler,
			data: {
				slideMasters: [
					{
						path: 'master-1',
						name: 'Corporate',
						elements: [],
						layouts: [
							{ path: 'layout-1', name: 'Title', elements: [] },
							{ path: 'layout-2', name: 'Custom Layout', elements: [] },
						],
					},
				],
			},
			target,
		});

		await run('addLayout');

		expect(applyMasterViewCrudAction).toHaveBeenCalledWith(
			handler,
			expect.anything(),
			'addLayout',
			{ tab: 'slides', masterIndex: 0, layoutIndex: null },
			{ name: undefined },
		);
		// The previous handler is kept (`keepPrevious`): the slides and the undo stack still show its
		// pictures, and disposing it revokes their Blob URLs.
		expect(setHandler).toHaveBeenCalledExactlyOnceWith(newHandler, true);
		expect(store.get().masterViewTarget).toStrictEqual({ masterIndex: 0, layoutIndex: 1 });
		expect(store.get().slideMasters[0]?.layouts).toHaveLength(2);
		expect(store.get().dirty).toBeTruthy();
	});

	it('prompts for a name on rename and cancels without calling apply when dismissed', async () => {
		const { run } = harness();
		stubActionList('renameMaster');

		const running = run('renameMaster');
		expect(await answerDialog('cancel')).toBe('New name');
		await running;

		expect(applyMasterViewCrudAction).not.toHaveBeenCalled();
	});

	it('passes the entered name through to applyMasterViewCrudAction', async () => {
		const { run, handler } = harness();
		stubActionList('renameMaster');
		applyMasterViewCrudAction.mockResolvedValue({ ok: false, reason: 'notFound' });

		const running = run('renameMaster');
		const openRoot = await openDialog();
		expect((openRoot.querySelector('input') as HTMLInputElement).value).toBe('Corporate');
		await answerDialog('confirm', 'Renamed Master');
		// The rejection notice follows; dismiss it.
		await answerDialog('confirm');
		await running;

		expect(applyMasterViewCrudAction).toHaveBeenCalledWith(
			handler,
			expect.anything(),
			'renameMaster',
			{ tab: 'slides', masterIndex: 0, layoutIndex: null },
			{ name: 'Renamed Master' },
		);
	});

	it('alerts the mapped reason on rejection and leaves the store untouched', async () => {
		const { run, store, setHandler } = harness();
		stubActionList('deleteMaster');
		applyMasterViewCrudAction.mockResolvedValue({ ok: false, reason: 'lastMaster' });
		const before = store.get().slideMasters;

		const running = run('deleteMaster');
		expect(await answerDialog('confirm')).toBe('The last slide master cannot be deleted.');
		await running;

		expect(setHandler).not.toHaveBeenCalled();
		expect(store.get().slideMasters).toBe(before);
	});

	it('alerts the notFound copy when the target vanished under the click', async () => {
		const { run } = harness();
		stubActionList('deleteLayout');
		applyMasterViewCrudAction.mockResolvedValue({ ok: false, reason: 'notFound' });

		const running = run('deleteLayout');
		expect(await answerDialog('confirm')).not.toBe('');
		await running;
	});

	it('leaves no undo or redo step that brings a deleted layout back', async () => {
		const home = 'layout-1';
		const spare = 'layout-2';
		const master = { path: 'master-1', name: 'Corporate', elements: [] };
		const { run, store, ops } = harness({
			slideMasters: [
				{
					...master,
					layouts: [
						{ path: home, name: 'Title', elements: [] },
						{ path: spare, name: 'Blank', elements: [] },
					],
				},
			],
			slides: [{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [], layoutPath: home }],
			masterViewTarget: { masterIndex: 0, layoutIndex: 1 },
		});
		const edit = (label: string, patch: Partial<PptxSlide>) => {
			ops.pushHistory(label);
			store.set({ slides: [{ ...store.get().slides[0]!, ...patch }] });
			ops.commitChange();
		};
		const reachesSpare = () =>
			store.get().slides.some((slide) => slide.layoutPath === spare) ||
			store
				.get()
				.slideMasters.some((entry) =>
					(entry.layouts ?? []).some((layout) => layout.path === spare),
				);
		// The spare layout is in an undo step (before 'two') and, once 'four' is undone, in the
		// redo step, but no slide uses it now.
		edit('one', { layoutPath: spare });
		edit('two', { layoutPath: home });
		edit('three', { backgroundColor: '#FF0000' });
		edit('four', { layoutPath: spare });
		ops.undo();
		expect(store.get().slides[0]?.layoutPath).toBe(home);
		stubActionList('deleteLayout');
		applyMasterViewCrudAction.mockResolvedValue({
			ok: true,
			handler: { getImageData: vi.fn() } as unknown as PptxHandler,
			data: {
				slideMasters: [{ ...master, layouts: [{ path: home, name: 'Title', elements: [] }] }],
			},
			target: { tab: 'slides', masterIndex: 0, layoutIndex: null },
		});

		await run('deleteLayout');

		expect(reachesSpare()).toBeFalsy();
		const reached: boolean[] = [];
		for (let step = 0; step < 6; step++) {
			ops.undo();
			reached.push(reachesSpare());
		}
		for (let step = 0; step < 6; step++) {
			ops.redo();
			reached.push(reachesSpare());
		}
		expect(reached).not.toContain(true);
		// Every step here also restores the masters, and they all still hold the spare layout.
		expect([ops.canUndo(), ops.canRedo()]).toStrictEqual([false, false]);
	});
});
