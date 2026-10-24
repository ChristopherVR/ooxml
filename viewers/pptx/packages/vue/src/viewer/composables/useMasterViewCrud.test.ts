// oxlint-disable react-hooks/rules-of-hooks
import { mount } from '@vue/test-utils';
import type {
	MasterLayoutCrudFailure,
	MasterViewCrudApplySuccess,
	MasterViewTarget,
} from 'ooxml-ui/pptx';
import type {
	PptxData,
	PptxHandler,
	PptxLayoutOption,
	PptxSlide,
	PptxSlideMaster,
} from 'pptx-viewer-core';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { defineComponent, h, shallowRef } from 'vue';

import type { EditorHistoryResult } from './useEditorHistory';
import { useEditorHistory } from './useEditorHistory';
import type { UseMasterViewCrudResult } from './useMasterViewCrud';

const { applyMasterViewCrudActionMock } = vi.hoisted(() => ({
	applyMasterViewCrudActionMock: vi.fn(),
}));

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	applyMasterViewCrudAction: applyMasterViewCrudActionMock,
}));

const { useMasterViewCrud } = await import('./useMasterViewCrud');

function master(path: string, overrides: Partial<PptxSlideMaster> = {}): PptxSlideMaster {
	return { path, name: 'Office Theme', layouts: [], ...overrides };
}

function baseData(masters: PptxSlideMaster[], slides: PptxData['slides'] = []): PptxData {
	return { slides, slideMasters: masters } satisfies Partial<PptxData> as PptxData;
}

/**
 * `useMasterViewCrud` calls `useI18n()` (rename prompt copy + failure
 * messages), so it has to run inside a component setup, exactly like
 * `useCustomShowsWiring` (see `authored-custom-show.test.ts`).
 */
function setup(
	masters: PptxSlideMaster[],
	target: MasterViewTarget,
	slides: PptxData['slides'] = [],
	handlerValue: PptxHandler | null = {} as PptxHandler,
): {
	crud: UseMasterViewCrudResult;
	masterRef: ReturnType<typeof shallowRef<PptxSlideMaster[]>>;
	layoutOptionsRef: ReturnType<typeof shallowRef<PptxLayoutOption[]>>;
	handler: ReturnType<typeof shallowRef<PptxHandler | null>>;
	onSelectMaster: ReturnType<typeof vi.fn>;
	onSelectLayout: ReturnType<typeof vi.fn>;
	markDirty: ReturnType<typeof vi.fn>;
	pushHistory: ReturnType<typeof vi.fn>;
} {
	const masterRef = shallowRef<PptxSlideMaster[]>(masters);
	const layoutOptionsRef = shallowRef<PptxLayoutOption[]>([]);
	const handler = shallowRef<PptxHandler | null>(handlerValue);
	const onSelectMaster = vi.fn();
	const onSelectLayout = vi.fn();
	const markDirty = vi.fn();
	const pushHistory = vi.fn();
	let captured: UseMasterViewCrudResult | null = null;
	mount(
		defineComponent({
			setup() {
				captured = useMasterViewCrud({
					handler,
					slideMasters: masterRef,
					layoutOptions: layoutOptionsRef,
					deckData: () => baseData(masterRef.value, slides),
					target: () => target,
					onSelectMaster,
					onSelectLayout,
					markDirty,
					pushHistory,
					forgetHistory: vi.fn(),
				});
				return () => h('div');
			},
		}),
	);
	return {
		crud: captured as unknown as UseMasterViewCrudResult,
		masterRef,
		layoutOptionsRef,
		handler,
		onSelectMaster,
		onSelectLayout,
		markDirty,
		pushHistory,
	};
}

/**
 * useMasterViewCrud: the wiring behind the Slide Master view sidebar's
 * Insert/Duplicate/Delete/Rename Layout and Slide Master buttons (wave-4 B4).
 * `applyMasterViewCrudAction` performs real ZIP surgery in `pptx-viewer-core`,
 * so it is mocked here; `masterViewCrudActions` (the button list) is pure and
 * runs for real.
 */
describe('useMasterViewCrud', () => {
	beforeEach(() => {
		applyMasterViewCrudActionMock.mockReset();
	});

	it('disables deleteLayout for a layout a slide still uses', () => {
		const layoutPath = 'ppt/slideLayouts/slideLayout1.xml';
		const { crud } = setup(
			[
				master('ppt/slideMasters/slideMaster1.xml', {
					layouts: [{ path: layoutPath, name: 'Title Slide' }],
				}),
			],
			{ tab: 'slides', masterIndex: 0, layoutIndex: 0 },
			[{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [], layoutPath }],
		);

		const deleteLayout = crud.actions.value.find((a) => a.id === 'deleteLayout');
		expect(deleteLayout?.enabled).toBeFalsy();
		expect(deleteLayout?.disabledReasonKey).toBe('pptx.masterView.layoutInUse');
		const addLayout = crud.actions.value.find((a) => a.id === 'addLayout');
		expect(addLayout?.enabled).toBeTruthy();
	});

	it('adopts the returned handler, slideMasters, layout list and selection when addLayout succeeds', async () => {
		const { crud, handler, masterRef, layoutOptionsRef, onSelectLayout, markDirty, pushHistory } =
			setup([master('ppt/slideMasters/slideMaster1.xml')], {
				tab: 'slides',
				masterIndex: 0,
				layoutIndex: null,
			});

		const newHandler = {} as PptxHandler;
		const newMasters: PptxSlideMaster[] = [
			master('ppt/slideMasters/slideMaster1.xml', {
				layouts: [{ path: 'ppt/slideLayouts/slideLayout9.xml', name: 'Title and Content' }],
			}),
		];
		// Home > Layout reads this list, kept apart from the masters it is built from: a command
		// that left it as loaded would keep offering a deleted layout and miss an inserted one.
		const newLayoutOptions: PptxLayoutOption[] = [
			{ path: 'ppt/slideLayouts/slideLayout9.xml', name: 'Title and Content' },
		];
		const success: MasterViewCrudApplySuccess = {
			ok: true,
			handler: newHandler,
			data: { ...baseData(newMasters), layoutOptions: newLayoutOptions },
			target: { tab: 'slides', masterIndex: 0, layoutIndex: 0 },
		};
		applyMasterViewCrudActionMock.mockResolvedValue(success);

		await crud.run('addLayout');

		expect(pushHistory).toHaveBeenCalledOnce();
		expect(handler.value).toBe(newHandler);
		expect(masterRef.value).toBe(newMasters);
		expect(layoutOptionsRef.value).toBe(newLayoutOptions);
		expect(onSelectLayout).toHaveBeenCalledWith(0, 0);
		expect(markDirty).toHaveBeenCalledOnce();
		expect(crud.error.value).toBeNull();
	});

	it('reports a failure without adopting anything', async () => {
		const { crud, handler } = setup(
			[
				master('ppt/slideMasters/slideMaster1.xml', {
					layouts: [{ path: 'ppt/slideLayouts/slideLayout1.xml', name: 'Title Slide' }],
				}),
			],
			{ tab: 'slides', masterIndex: 0, layoutIndex: 0 },
		);
		const originalHandler = handler.value;
		const failure: MasterLayoutCrudFailure = { ok: false, reason: 'inUse' };
		applyMasterViewCrudActionMock.mockResolvedValue(failure);

		await crud.run('deleteLayout');

		expect(handler.value).toBe(originalHandler);
		expect(crud.error.value).toBe('This layout is used by slides and cannot be deleted.');
	});

	it('prompts for a name before renaming, and skips the call when the prompt is cancelled', async () => {
		// happy-dom does not implement `window.prompt` at all, so there is no
		// existing method for `vi.spyOn` to wrap; `vi.stubGlobal` installs one.
		const promptSpy = vi.fn().mockReturnValue(null);
		vi.stubGlobal('prompt', promptSpy);
		const { crud } = setup([master('ppt/slideMasters/slideMaster1.xml')], {
			tab: 'slides',
			masterIndex: 0,
			layoutIndex: null,
		});

		await crud.run('renameMaster');

		expect(promptSpy).toHaveBeenCalledWith('New name', 'Office Theme');
		expect(applyMasterViewCrudActionMock).not.toHaveBeenCalled();
		vi.unstubAllGlobals();
	});

	it('leaves no undo or redo step that puts a slide back on a deleted layout', async () => {
		const masterPath = 'ppt/slideMasters/slideMaster1.xml';
		const home = 'ppt/slideLayouts/slideLayout1.xml';
		const spare = 'ppt/slideLayouts/slideLayout2.xml';
		const masterRef = shallowRef<PptxSlideMaster[]>([
			master(masterPath, {
				layouts: [
					{ path: home, name: 'Title Slide' },
					{ path: spare, name: 'Blank' },
				],
			}),
		]);
		const slides = shallowRef<PptxSlide[]>([
			{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [], layoutPath: home },
		]);
		let history: EditorHistoryResult | null = null;
		let crud: UseMasterViewCrudResult | null = null;
		mount(
			defineComponent({
				setup() {
					history = useEditorHistory(slides);
					crud = useMasterViewCrud({
						handler: shallowRef<PptxHandler | null>({} as PptxHandler),
						slideMasters: masterRef,
						layoutOptions: shallowRef<PptxLayoutOption[]>([]),
						deckData: () => baseData(masterRef.value, slides.value),
						target: () => ({ tab: 'slides', masterIndex: 0, layoutIndex: 1 }),
						onSelectMaster: vi.fn(),
						onSelectLayout: vi.fn(),
						markDirty: vi.fn(),
						pushHistory: history.pushHistory,
						forgetHistory: history.forgetHistory,
					});
					return () => h('div');
				},
			}),
		);
		const undoStack = history as unknown as EditorHistoryResult;
		const edit = (label: string, patch: Partial<PptxSlide>) => {
			undoStack.pushHistory(label);
			slides.value = [{ ...slides.value[0]!, ...patch }];
		};
		const layoutsInUse = () => slides.value.map((slide) => slide.layoutPath);
		// The spare layout is in an undo step (before 'two') and, once 'four' is undone, in the
		// redo step, but no slide uses it now.
		edit('one', { layoutPath: spare });
		edit('two', { layoutPath: home });
		edit('three', { backgroundColor: '#FF0000' });
		edit('four', { layoutPath: spare });
		undoStack.undo();
		expect(layoutsInUse()).toStrictEqual([home]);
		applyMasterViewCrudActionMock.mockResolvedValue({
			ok: true,
			handler: {} as PptxHandler,
			data: baseData([master(masterPath, { layouts: [{ path: home, name: 'Title Slide' }] })]),
			target: { tab: 'slides', masterIndex: 0, layoutIndex: null },
		} satisfies MasterViewCrudApplySuccess);

		await (crud as unknown as UseMasterViewCrudResult).run('deleteLayout');

		// Left: the command's own step and the one before 'three'; the command emptied the redo stack.
		expect(undoStack.canUndo.value).toBeTruthy();
		expect(undoStack.canRedo.value).toBeFalsy();
		const reached: (string | undefined)[] = [];
		let undos = 0;
		while (undoStack.canUndo.value) {
			undoStack.undo();
			undos += 1;
			reached.push(...layoutsInUse());
		}
		while (undoStack.canRedo.value) {
			undoStack.redo();
			reached.push(...layoutsInUse());
		}
		expect(reached).not.toContain(spare);
		expect(undos).toBe(2);
	});
});
