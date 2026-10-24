import { PresentationBuilder } from 'ooxml-core/pptx';
import type { PptxData, PptxSlide, PptxSlideMaster } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createEditorHistoryController } from './controller/history-controller';
import type { MasterViewTarget } from './master-view';
import { applyMasterViewCrudAction } from './master-view-crud';
import { historyStepUsesParts, removedSlideMasterParts } from './master-view-crud-history';

const ON_MASTER: MasterViewTarget = { tab: 'slides', masterIndex: 0, layoutIndex: null };

function layoutTarget(data: PptxData, name: string): MasterViewTarget {
	const layoutIndex = data.slideMasters![0]!.layouts!.findIndex((layout) => layout.name === name);
	expect(layoutIndex).toBeGreaterThanOrEqual(0);
	return { tab: 'slides', masterIndex: 0, layoutIndex };
}

function onLayout(layoutPath: string): Pick<PptxSlide, 'layoutPath'> {
	return { layoutPath };
}

describe('removedSlideMasterParts', () => {
	it('names the layout Delete Layout removed, and nothing after a command that adds', async () => {
		const { handler, data } = await PresentationBuilder.create();
		const blank = layoutTarget(data, 'Blank');
		const blankPath = data.slideMasters![0]!.layouts![blank.layoutIndex!]!.path;

		const deleted = await applyMasterViewCrudAction(handler, data, 'deleteLayout', blank);
		const added = await applyMasterViewCrudAction(handler, data, 'addLayout', ON_MASTER);

		expect(deleted.ok && added.ok).toBeTruthy();
		if (!deleted.ok || !added.ok) {
			return;
		}
		expect(removedSlideMasterParts(data, deleted.data)).toStrictEqual([blankPath]);
		expect(removedSlideMasterParts(data, added.data)).toStrictEqual([]);
	});

	it('names a deleted master with every one of its layouts', async () => {
		const { handler, data } = await PresentationBuilder.create();
		const added = await applyMasterViewCrudAction(handler, data, 'addMaster', ON_MASTER);
		expect(added.ok).toBeTruthy();
		if (!added.ok) {
			return;
		}
		const second = added.data.slideMasters![added.target.masterIndex]!;
		const deleted = await applyMasterViewCrudAction(
			added.handler,
			added.data,
			'deleteMaster',
			added.target,
		);
		expect(deleted.ok).toBeTruthy();
		if (!deleted.ok) {
			return;
		}

		expect(removedSlideMasterParts(added.data, deleted.data)).toStrictEqual([
			second.path,
			...(second.layouts ?? []).map((layout) => layout.path),
		]);
	});
});

describe('historyStepUsesParts', () => {
	const master: PptxSlideMaster = {
		path: 'ppt/slideMasters/slideMaster1.xml',
		layouts: [{ path: 'ppt/slideLayouts/slideLayout2.xml', name: 'Two' }],
	} as PptxSlideMaster;

	it('matches a step with a slide on a removed layout', () => {
		const removed = ['ppt/slideLayouts/slideLayout2.xml'];
		expect(historyStepUsesParts({ slides: [{}, onLayout(removed[0]!)] }, removed)).toBeTruthy();
		expect(
			historyStepUsesParts({ slides: [onLayout('ppt/slideLayouts/slideLayout3.xml')] }, removed),
		).toBeFalsy();
	});

	it('matches a step whose masters still hold a removed master or layout', () => {
		expect(
			historyStepUsesParts({ slides: [], slideMasters: [master] }, [
				'ppt/slideLayouts/slideLayout2.xml',
			]),
		).toBeTruthy();
		expect(
			historyStepUsesParts({ slides: [], slideMasters: [{ ...master, layouts: [] }] }, [
				master.path,
			]),
		).toBeTruthy();
		expect(
			historyStepUsesParts({ slides: [], slideMasters: [master] }, [
				'ppt/slideLayouts/slideLayout9.xml',
			]),
		).toBeFalsy();
	});

	it('matches nothing when the command removed nothing', () => {
		expect(
			historyStepUsesParts({ slides: [onLayout('x')], slideMasters: [master] }, []),
		).toBeFalsy();
	});

	it('leaves no undo or redo step that puts a slide back on a deleted layout', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(createSlide('Title Slide').build());
		const blank = layoutTarget(data, 'Blank');
		const blankPath = data.slideMasters![0]!.layouts![blank.layoutIndex!]!.path;
		const titlePath = data.slides[0]!.layoutPath!;
		const deck = (...layouts: string[]) => ({ slides: layouts.map(onLayout) });
		const history = createEditorHistoryController<{ slides: Pick<PptxSlide, 'layoutPath'>[] }>();
		// Five steps, two of them undone: a slide is on Blank in an undo step (before two) and in a
		// redo step (behind four), but not in the deck as it is now.
		history.record(deck(titlePath), 'one');
		history.record(deck(blankPath), 'two');
		history.record(deck(titlePath, titlePath), 'three');
		history.record(deck(titlePath, titlePath, titlePath), 'four');
		history.record(deck(titlePath, titlePath, titlePath, titlePath), 'five');
		history.undo(deck(blankPath, titlePath, titlePath, titlePath));
		history.undo(deck(titlePath, titlePath, titlePath, titlePath));

		const result = await applyMasterViewCrudAction(handler, data, 'deleteLayout', blank);
		expect(result.ok).toBeTruthy();
		if (!result.ok) {
			return;
		}
		const removed = removedSlideMasterParts(data, result.data);
		history.forget((step) => historyStepUsesParts(step, removed));

		expect(history.getSnapshot()).toMatchObject({
			canUndo: true,
			canRedo: true,
			depth: 1,
			undoLabel: 'three',
			redoLabel: 'four',
		});
	});
});
