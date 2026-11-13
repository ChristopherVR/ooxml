import type { PptxCustomShow, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { BuildingBlockAvailabilityInput } from './building-block-actions';
import {
	applyCustomShowEdit,
	describeCustomShows,
	resolveCustomShowStart,
} from './building-block-custom-shows';
import {
	describeBroadcastSupport,
	describeShortcutCatalogue,
	describeShowDialogActions,
} from './building-block-show-dialogs';
import {
	applySlideShowSettingsPatch,
	describeSlideShowSettings,
	resolveSlideShowSettings,
} from './building-block-show-settings';

const ready: BuildingBlockAvailabilityInput = {
	ready: true,
	hasActiveSlide: true,
	canEdit: true,
	editing: true,
	canInsert: true,
	hasSelection: false,
	selectionHasHyperlink: false,
	selectionHasEquation: false,
	layoutCount: 0,
	hasCurrentLayout: false,
	hasHandler: true,
};
const slide = (rId: string): PptxSlide => ({ id: rId, rId, elements: [] }) as unknown as PptxSlide;
const slides = [slide('rId1'), slide('rId2'), slide('rId3')];
const shows: PptxCustomShow[] = [{ id: '0', name: 'Short', slideRIds: ['rId1', 'rId3'] }];
const context = { slideCount: 3, customShows: shows };

describe('slide show settings', () => {
	it('applies the dialog rules', () => {
		expect(applySlideShowSettingsPatch({}, { showType: 'kiosk' }, context).loopContinuously).toBe(
			true,
		);
		const range = applySlideShowSettingsPatch({}, { showSlidesMode: 'range' }, context);
		expect([range.showSlidesFrom, range.showSlidesTo]).toEqual([1, 3]);
		expect(
			applySlideShowSettingsPatch({}, { showSlidesMode: 'customShow' }, context)
				.showSlidesCustomShowId,
		).toBe('0');
	});

	it('refuses bad ranges, unknown shows and colours with codes', () => {
		const code = (patch: Parameters<typeof resolveSlideShowSettings>[1]) => {
			const r = resolveSlideShowSettings({}, patch, context);
			return r.ok ? 'ok' : r.code;
		};
		expect(code({ showSlidesMode: 'range', showSlidesFrom: 3, showSlidesTo: 2 })).toBe(
			'invalid-range',
		);
		expect(code({ showSlidesMode: 'range', showSlidesTo: 9 })).toBe('invalid-range');
		expect(code({ showSlidesMode: 'customShow', showSlidesCustomShowId: 'nope' })).toBe(
			'unknown-show',
		);
		expect(code({ penColor: 'red' })).toBe('invalid-data');
		expect(code({ showType: 'browsed', penColor: '#FF0000' })).toBe('ok');
	});

	it('describes defaults and inverts the stored negations', () => {
		const view = describeSlideShowSettings({ showWithNarration: false }, context);
		expect(view).toMatchObject({
			showType: 'presented',
			playNarration: false,
			playAnimation: true,
			advanceMode: 'useTimings',
			slidesMode: 'all',
			rangeTo: 3,
			showScrollbar: true,
		});
	});
});

describe('custom show edits', () => {
	it('creates, renames, sets slides, copies and removes', () => {
		const created = applyCustomShowEdit(shows, slides, {
			kind: 'create',
			name: 'Short',
			slideRIds: ['rId2'],
		});
		expect(created.ok && created.shows.map((s) => [s.id, s.name, s.slideRIds])).toEqual([
			['0', 'Short', ['rId1', 'rId3']],
			['1', 'Short (2)', ['rId2']],
		]);
		const renamed = applyCustomShowEdit(shows, slides, { kind: 'rename', id: '0', name: 'Intro' });
		expect(renamed.ok && renamed.shows[0]?.name).toBe('Intro');
		const set = applyCustomShowEdit(shows, slides, {
			kind: 'setSlides',
			id: '0',
			slideRIds: ['rId2'],
		});
		expect(set.ok && set.shows[0]?.slideRIds).toEqual(['rId2']);
		const copied = applyCustomShowEdit(shows, slides, { kind: 'copy', id: '0' });
		expect(copied.ok && copied.shows[1]?.name).toBe('Copy of Short');
		const removed = applyCustomShowEdit(shows, slides, { kind: 'remove', id: '0' });
		expect(removed.ok && removed.shows).toEqual([]);
	});

	it('refuses unknown shows and slides and empty names', () => {
		const code = (edit: Parameters<typeof applyCustomShowEdit>[2]) => {
			const r = applyCustomShowEdit(shows, slides, edit);
			return r.ok ? 'ok' : r.code;
		};
		expect(code({ kind: 'rename', id: '9', name: 'x' })).toBe('unknown-show');
		expect(code({ kind: 'setSlides', id: '0', slideRIds: ['zzz'] })).toBe('unknown-slide');
		expect(code({ kind: 'create', name: '  ' })).toBe('invalid-name');
	});

	it('resolves a start and describes the shows', () => {
		expect(resolveCustomShowStart(shows, slides, '0')).toEqual({ ok: true, firstSlideIndex: 0 });
		expect(resolveCustomShowStart(shows, slides, '5')).toEqual({ ok: false, code: 'unknown-show' });
		expect(resolveCustomShowStart([{ id: '1', name: 'e', slideRIds: [] }], slides, '1')).toEqual({
			ok: false,
			code: 'no-slide',
		});
		expect(describeCustomShows(shows, slides).shows[0]?.slideCount).toBe(2);
	});
});

describe('show dialog availability and catalogue', () => {
	it('gates the editing dialogs and reports broadcast as unsupported', () => {
		const ok = describeShowDialogActions(ready);
		expect(ok.setUpSlideShow).toEqual({ available: true, needsHostUi: true });
		expect(ok.broadcast).toEqual({ available: false, reason: 'unsupported', needsHostUi: true });
		const readOnly = describeShowDialogActions({ ...ready, canEdit: false });
		expect(readOnly.customShows.reason).toBe('read-only');
		expect(readOnly.shortcuts.available).toBe(true);
		expect(describeBroadcastSupport().supported).toBe(false);
	});

	it('lists the shortcuts as unique localised data', () => {
		const catalogue = describeShortcutCatalogue((key) => `T:${key}`);
		expect(catalogue.length).toBeGreaterThan(20);
		expect(new Set(catalogue.map((entry) => entry.id)).size).toBe(catalogue.length);
		const redo = catalogue.find((entry) => entry.id === 'toolbar.redo');
		expect(redo?.alternatives).toEqual(['Ctrl/Cmd+Shift+Z', 'Ctrl/Cmd+Y']);
		expect(redo?.description).toBe('T:pptx.toolbar.redo');
	});
});
