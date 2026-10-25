// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';

import { createCustomizationController } from '../customization/customization-controller';
import { customizeOptionsTabs } from '../customization/customization-surfaces';
import {
	EXTERNAL_CONTENT_SETTING_ID,
	lockExternalContentSetting,
	resolveAllowExternalContent,
} from './external-content-policy';
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
