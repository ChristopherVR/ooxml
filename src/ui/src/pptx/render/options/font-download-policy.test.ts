// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createCustomizationController } from '../customization/customization-controller';
import { customizeOptionsTabs } from '../customization/customization-surfaces';
import { GOOGLE_FONTS_CSS2_BASE } from '../google-webfonts';
import { EXTERNAL_CONTENT_SETTING_ID } from './external-content-policy';
import {
	FONT_DOWNLOAD_SETTING_ID,
	hostFontDownloadLock,
	lockFontDownloadSetting,
	resolveEffectiveGoogleFonts,
	resolveFontDownloadBase,
} from './font-download-policy';
import { createViewerOptionsStore } from './viewer-options-store';
import { VIEWER_OPTIONS_TABS } from './viewer-options-schema';
import { DEFAULT_VIEWER_OPTIONS } from './viewer-options';

const trust = (allowFontDownloads: boolean) => ({ trust: { allowFontDownloads } }) as never;
const MIRROR = 'https://fonts.example.test/css2';
const INVALID = { baseUrl: 'ftp://fonts.example.test/css2' };

describe('font download policy', () => {
	beforeEach(() => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
	});
	afterEach(() => {
		vi.restoreAllMocks();
	});

	it('defaults the Trust Center option on', () => {
		expect(DEFAULT_VIEWER_OPTIONS.trust.allowFontDownloads).toBeTruthy();
	});

	it('lets the Trust Center option decide when the host leaves googleFonts unset', () => {
		expect(resolveFontDownloadBase(undefined, trust(true))).toBe(GOOGLE_FONTS_CSS2_BASE);
		expect(resolveFontDownloadBase(undefined, trust(false))).toBeNull();
		expect(resolveEffectiveGoogleFonts(undefined, trust(true))).toBeUndefined();
		expect(resolveEffectiveGoogleFonts(undefined, trust(false))).toBe(false);
	});

	it('uses the host value over the Trust Center option', () => {
		expect(resolveFontDownloadBase(false, trust(true))).toBeNull();
		expect(resolveFontDownloadBase({ baseUrl: MIRROR }, trust(false))).toBe(MIRROR);
		expect(resolveFontDownloadBase(INVALID, trust(true))).toBeNull();
	});

	it('locks the toggle to what the host value allows', () => {
		expect(hostFontDownloadLock(undefined)).toBeUndefined();
		expect(hostFontDownloadLock(false)).toBe(false);
		expect(hostFontDownloadLock({ baseUrl: MIRROR })).toBe(true);
		expect(hostFontDownloadLock(INVALID)).toBe(false);
		const locked = { 'general.userName': 'Kiosk' } as const;
		expect(lockFontDownloadSetting(locked, undefined)).toBe(locked);
		expect(lockFontDownloadSetting(locked, false)).toStrictEqual({
			'general.userName': 'Kiosk',
			[FONT_DOWNLOAD_SETTING_ID]: false,
		});
		expect(
			lockFontDownloadSetting({ [FONT_DOWNLOAD_SETTING_ID]: false }, { baseUrl: MIRROR })[
				FONT_DOWNLOAD_SETTING_ID
			],
		).toBe(true);
	});
});

describe('customization controller googleFonts', () => {
	beforeEach(() => {
		localStorage.clear();
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
	});
	afterEach(() => {
		vi.restoreAllMocks();
	});

	it('locks, re-applies and releases the Trust setting', () => {
		const controller = createCustomizationController();
		let calls = 0;
		controller.subscribe(() => calls++);
		controller.setGoogleFonts(false);
		expect(controller.getResolved().lockedSettings[FONT_DOWNLOAD_SETTING_ID]).toBe(false);
		controller.setGoogleFonts(INVALID);
		expect(calls).toBe(1);
		controller.setGoogleFonts({ baseUrl: MIRROR });
		controller.setGoogleFonts({ baseUrl: MIRROR });
		expect(controller.getResolved().lockedSettings[FONT_DOWNLOAD_SETTING_ID]).toBe(true);
		expect(calls).toBe(2);
		controller.setGoogleFonts(undefined);
		expect(
			Object.hasOwn(controller.getResolved().lockedSettings, FONT_DOWNLOAD_SETTING_ID),
		).toBeFalsy();
		expect(calls).toBe(3);
	});

	it('keeps both host locks across customization edits and resets', () => {
		const controller = createCustomizationController();
		controller.setGoogleFonts(false);
		controller.setAllowExternalContent(true);
		controller.api.lockSetting('general.userName', 'Kiosk');
		controller.api.resetCustomization();
		expect(controller.getResolved().lockedSettings).toStrictEqual({
			[EXTERNAL_CONTENT_SETTING_ID]: true,
			[FONT_DOWNLOAD_SETTING_ID]: false,
		});
		controller.setAllowExternalContent(undefined);
		expect(controller.getResolved().lockedSettings).toStrictEqual({
			[FONT_DOWNLOAD_SETTING_ID]: false,
		});
	});

	it('forces the option value and renders the Trust toggle read-only', () => {
		const controller = createCustomizationController();
		controller.setGoogleFonts(false);
		const store = createViewerOptionsStore();
		store.setConstraints({ locked: controller.getResolved().lockedSettings });
		expect(store.getOptions().trust.allowFontDownloads).toBeFalsy();
		store.setValue('trust', 'allowFontDownloads', true);
		expect(store.getOptions().trust.allowFontDownloads).toBeFalsy();
		const control = customizeOptionsTabs(VIEWER_OPTIONS_TABS, controller.getResolved())
			.flatMap((tab) => tab.sections)
			.flatMap((section) => section.controls)
			.find((c) => c.group === 'trust' && c.key === 'allowFontDownloads');
		expect(control?.readOnly).toBeTruthy();
	});
});
