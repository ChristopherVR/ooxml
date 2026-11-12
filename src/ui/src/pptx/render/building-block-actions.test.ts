import { describe, expect, it } from 'vitest';

import {
	describeBuildingBlockActions,
	HOST_IMAGE_MAX_BYTES,
	validateHostImageFile,
	validateHostMediaFile,
} from './building-block-actions';
import type { BuildingBlockAvailabilityInput } from './building-block-actions';

const ready: BuildingBlockAvailabilityInput = {
	ready: true,
	hasActiveSlide: true,
	canEdit: true,
	editing: true,
	canInsert: true,
	hasSelection: true,
	selectionHasHyperlink: true,
	selectionHasEquation: true,
	layoutCount: 3,
	hasCurrentLayout: true,
	hasHandler: true,
};

describe('validateHostImageFile', () => {
	it('accepts a non-empty image within the limit', () => {
		expect(validateHostImageFile({ type: 'image/png', size: 10 })).toStrictEqual({ ok: true });
	});

	it('names why a file is refused', () => {
		expect(validateHostImageFile(null)).toStrictEqual({ ok: false, code: 'no-file' });
		expect(validateHostImageFile({ type: 'application/pdf', size: 10 })).toStrictEqual({
			ok: false,
			code: 'unsupported-type',
		});
		expect(validateHostImageFile({ type: 'image/png', size: 0 })).toStrictEqual({
			ok: false,
			code: 'empty-file',
		});
		expect(
			validateHostImageFile({ type: 'image/png', size: HOST_IMAGE_MAX_BYTES + 1 }),
		).toStrictEqual({ ok: false, code: 'file-too-large' });
		expect(validateHostImageFile({ type: 'image/png', size: 11 }, 10)).toStrictEqual({
			ok: false,
			code: 'file-too-large',
		});
	});
});

describe('validateHostMediaFile', () => {
	it('reports the kind of an accepted file', () => {
		expect(validateHostMediaFile({ type: 'audio/mpeg', size: 1 })).toStrictEqual({
			ok: true,
			mediaType: 'audio',
		});
		expect(validateHostMediaFile({ type: 'video/mp4', size: 1 })).toStrictEqual({
			ok: true,
			mediaType: 'video',
		});
	});

	it('refuses other types, empty and oversized files', () => {
		expect(validateHostMediaFile({ type: 'image/png', size: 1 })).toStrictEqual({
			ok: false,
			code: 'unsupported-type',
		});
		expect(validateHostMediaFile({ type: 'video/mp4', size: 0 })).toStrictEqual({
			ok: false,
			code: 'empty-file',
		});
		expect(validateHostMediaFile({ type: 'video/mp4', size: 5 }, 4)).toStrictEqual({
			ok: false,
			code: 'file-too-large',
		});
	});
});

describe('describeBuildingBlockActions', () => {
	it('marks the picker and dialog workflows as needing host UI, and the rest as runnable', () => {
		const state = describeBuildingBlockActions(ready);
		expect(state.insertImage).toStrictEqual({ available: true, needsHostUi: true });
		expect(state.hyperlink.needsHostUi).toBeTruthy();
		expect(state.insertSmartArt.needsHostUi).toBeTruthy();
		expect(state.editEquation.needsHostUi).toBeTruthy();
		expect(state.applyLayout).toStrictEqual({ available: true, needsHostUi: false });
		expect(state.replace).toStrictEqual({ available: true, needsHostUi: false });
	});

	it('blocks every edit when the viewer is read-only, protected or not ready', () => {
		expect(describeBuildingBlockActions({ ...ready, canEdit: false }).insertImage.reason).toBe(
			'read-only',
		);
		expect(describeBuildingBlockActions({ ...ready, editing: false }).replace.reason).toBe(
			'read-only',
		);
		expect(describeBuildingBlockActions({ ...ready, canInsert: false }).insertMedia.reason).toBe(
			'protected',
		);
		expect(describeBuildingBlockActions({ ...ready, ready: false }).applyLayout.reason).toBe(
			'not-ready',
		);
	});

	it('needs a selection, a link or an equation where the action acts on one', () => {
		const none = describeBuildingBlockActions({
			...ready,
			hasSelection: false,
			selectionHasHyperlink: false,
			selectionHasEquation: false,
		});
		expect(none.hyperlink.reason).toBe('no-selection');
		expect(none.editEquation.reason).toBe('no-selection');
		const plain = describeBuildingBlockActions({
			...ready,
			selectionHasHyperlink: false,
			selectionHasEquation: false,
		});
		expect(plain.removeHyperlink.reason).toBe('no-hyperlink');
		expect(plain.editEquation.reason).toBe('no-equation');
		expect(plain.hyperlink.available).toBeTruthy();
	});

	it('reports layouts as unavailable without a handler, layouts or a recorded layout', () => {
		expect(describeBuildingBlockActions({ ...ready, layoutCount: 0 }).applyLayout.reason).toBe(
			'no-layout',
		);
		expect(describeBuildingBlockActions({ ...ready, hasHandler: false }).resetLayout.reason).toBe(
			'no-layout',
		);
		expect(
			describeBuildingBlockActions({ ...ready, hasCurrentLayout: false }).resetLayout.reason,
		).toBe('no-layout');
		expect(
			describeBuildingBlockActions({ ...ready, hasActiveSlide: false }).insertImage.reason,
		).toBe('no-slide');
	});
});
