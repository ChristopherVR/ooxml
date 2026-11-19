// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';

import { createCustomizationController } from '../customization/customization-controller';
import { customizeOptionsTabs } from '../customization/customization-surfaces';
import {
	dispatchExternalContentBarIntent,
	EXTERNAL_CONTENT_SETTING_ID,
	externalContentBarVisible,
	lockExternalContentSetting,
	prepareExternalContentReload,
	resolveAllowExternalContent,
} from './external-content-policy';
import { DEFAULT_VIEWER_OPTIONS } from './viewer-options';
import { createViewerOptionsStore } from './viewer-options-store';
import { VIEWER_OPTIONS_TABS } from './viewer-options-schema';

const trust = (allowExternalContent: boolean) => ({ trust: { allowExternalContent } }) as never;

describe('resolveAllowExternalContent', () => {
	it('uses the host value when set and the user option otherwise', () => {
		expect(resolveAllowExternalContent(false, trust(true))).toBeFalsy();
		expect(resolveAllowExternalContent(true, trust(false))).toBeTruthy();
		expect(resolveAllowExternalContent(undefined, trust(true))).toBeTruthy();
		expect(resolveAllowExternalContent(undefined, trust(false))).toBeFalsy();
	});

	it('blocks by default, like PowerPoint', () => {
		expect(DEFAULT_VIEWER_OPTIONS.trust.allowExternalContent).toBe(false);
		expect(resolveAllowExternalContent(undefined, DEFAULT_VIEWER_OPTIONS)).toBe(false);
	});

	it('allows the presentation the user enabled, unless the host blocks it', () => {
		expect(resolveAllowExternalContent(undefined, trust(false), true)).toBe(true);
		expect(resolveAllowExternalContent(false, trust(false), true)).toBe(false);
	});
});

describe('externalContentBarVisible', () => {
	const shown = {
		hasLinkedWebContent: true,
		allowed: false,
		host: undefined,
		closed: false,
	} as const;

	it('shows for a deck whose linked content the load left out', () => {
		expect(externalContentBarVisible(shown)).toBe(true);
	});

	it('hides without linked content, once allowed, when the host blocks or the user closes it', () => {
		expect(externalContentBarVisible({ ...shown, hasLinkedWebContent: undefined })).toBe(false);
		expect(externalContentBarVisible({ ...shown, allowed: true })).toBe(false);
		expect(externalContentBarVisible({ ...shown, host: false })).toBe(false);
		expect(externalContentBarVisible({ ...shown, closed: true })).toBe(false);
	});
});

describe('prepareExternalContentReload', () => {
	const current = new Uint8Array([1, 2, 3]);

	it('reloads the deck as it stands even with nothing unsaved', async () => {
		await expect(
			prepareExternalContentReload({
				serializeCurrentDeck: async () => current,
				isDirty: false,
				slideIndex: 2,
			}),
		).resolves.toStrictEqual({ bytes: current, keepDirty: false, slideIndex: 2 });
	});

	it('keeps unsaved changes unsaved', async () => {
		const reload = await prepareExternalContentReload({
			serializeCurrentDeck: async () => current,
			isDirty: true,
			slideIndex: 0,
		});
		expect(reload?.keepDirty).toBe(true);
	});

	it('reloads nothing without a deck', async () => {
		await expect(
			prepareExternalContentReload({
				serializeCurrentDeck: async () => null,
				isDirty: false,
				slideIndex: 0,
			}),
		).resolves.toBeNull();
	});
});

describe('dispatchExternalContentBarIntent', () => {
	it('routes Enable Content and Close', () => {
		const calls: string[] = [];
		const handlers = { enable: () => calls.push('enable'), close: () => calls.push('close') };
		dispatchExternalContentBarIntent(new CustomEvent('x', { detail: { id: 'action' } }), handlers);
		dispatchExternalContentBarIntent(new CustomEvent('x', { detail: { id: 'close' } }), handlers);
		dispatchExternalContentBarIntent(new CustomEvent('x'), handlers);
		expect(calls).toEqual(['enable', 'close']);
	});
});

describe('lockExternalContentSetting', () => {
	it('returns the same locks when unset and adds the Trust lock otherwise', () => {
		const locked = { 'general.userName': 'Kiosk' } as const;
		expect(lockExternalContentSetting(locked, undefined)).toBe(locked);
		expect(lockExternalContentSetting(locked, false)).toStrictEqual({
			'general.userName': 'Kiosk',
			[EXTERNAL_CONTENT_SETTING_ID]: false,
		});
		expect(
			lockExternalContentSetting({ [EXTERNAL_CONTENT_SETTING_ID]: false }, true)[
				EXTERNAL_CONTENT_SETTING_ID
			],
		).toBeTruthy();
	});
});

describe('customization controller allowExternalContent', () => {
	beforeEach(() => {
		localStorage.clear();
	});

	it('locks, re-applies and releases the Trust setting', () => {
		const controller = createCustomizationController();
		let calls = 0;
		controller.subscribe(() => calls++);
		expect(
			Object.hasOwn(controller.getResolved().lockedSettings, EXTERNAL_CONTENT_SETTING_ID),
		).toBeFalsy();
		controller.setAllowExternalContent(false);
		expect(controller.getResolved().lockedSettings[EXTERNAL_CONTENT_SETTING_ID]).toBe(false);
		controller.setAllowExternalContent(false);
		expect(calls).toBe(1);
		controller.setAllowExternalContent(true);
		expect(controller.getResolved().lockedSettings[EXTERNAL_CONTENT_SETTING_ID]).toBe(true);
		controller.setAllowExternalContent(undefined);
		expect(
			Object.hasOwn(controller.getResolved().lockedSettings, EXTERNAL_CONTENT_SETTING_ID),
		).toBeFalsy();
		expect(calls).toBe(3);
	});

	it('keeps the host switch across customization edits and resets', () => {
		const controller = createCustomizationController();
		controller.setAllowExternalContent(false);
		controller.api.lockSetting('general.userName', 'Kiosk');
		controller.api.resetCustomization();
		expect(controller.getResolved().lockedSettings).toStrictEqual({
			[EXTERNAL_CONTENT_SETTING_ID]: false,
		});
	});

	it('forces the option value and renders the Trust toggle read-only', () => {
		const controller = createCustomizationController();
		controller.setAllowExternalContent(false);
		const store = createViewerOptionsStore();
		store.setValue('trust', 'allowExternalContent', true);
		store.setConstraints({ locked: controller.getResolved().lockedSettings });
		expect(store.getOptions().trust.allowExternalContent).toBeFalsy();
		store.setValue('trust', 'allowExternalContent', true);
		expect(store.getOptions().trust.allowExternalContent).toBeFalsy();
		const control = customizeOptionsTabs(VIEWER_OPTIONS_TABS, controller.getResolved())
			.flatMap((tab) => tab.sections)
			.flatMap((section) => section.controls)
			.find((c) => c.group === 'trust' && c.key === 'allowExternalContent');
		expect(control?.readOnly).toBeTruthy();
	});
});
