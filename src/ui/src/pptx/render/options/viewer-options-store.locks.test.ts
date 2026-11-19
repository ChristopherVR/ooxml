// @vitest-environment happy-dom
/**
 * A host lock forces a setting for one viewer session; it must never replace
 * or delete the value the user saved, because the next viewer on the same
 * origin may run without the lock.
 */
import { beforeEach, describe, expect, it } from 'vitest';

import { VIEWER_PREFS_STORAGE_KEY } from '../viewer-prefs-storage';
import { createViewerOptionsStore } from './viewer-options-store';

const LOCK_OFF = { locked: { 'trust.allowExternalContent': false } };
const LOCK_ON = { locked: { 'trust.allowExternalContent': true } };

function storedTrust(): unknown {
	const raw = localStorage.getItem(VIEWER_PREFS_STORAGE_KEY);
	return raw ? (JSON.parse(raw) as { options?: { trust?: unknown } }).options?.trust : undefined;
}

function saveUserChoice(allowExternalContent: boolean): void {
	createViewerOptionsStore().setValue('trust', 'allowExternalContent', allowExternalContent);
	expect(storedTrust()).toStrictEqual({ allowExternalContent });
}

describe('viewer options store host locks', () => {
	beforeEach(() => {
		localStorage.clear();
	});

	it("keeps the user's saved value for a locked setting across another write", () => {
		saveUserChoice(true);
		const locked = createViewerOptionsStore();
		locked.setConstraints(LOCK_OFF);
		expect(locked.getOptions().trust.allowExternalContent).toBe(false);
		// Any later persisted write must not drop the user's choice for the locked key.
		locked.setValue('general', 'userName', 'Kiosk');
		expect(storedTrust()).toStrictEqual({ allowExternalContent: true });

		// The next viewer on this origin, without the lock, honours the user's choice.
		expect(createViewerOptionsStore().getOptions().trust.allowExternalContent).toBe(true);
	});

	it('persists no forced value as the user choice', () => {
		const locked = createViewerOptionsStore();
		locked.setConstraints(LOCK_ON);
		locked.setValue('general', 'userName', 'Kiosk');
		expect(storedTrust()).toBeUndefined();
		expect(createViewerOptionsStore().getOptions().trust.allowExternalContent).toBe(false);
	});

	it("restores the user's own value when the lock is lifted at runtime", () => {
		saveUserChoice(true);
		const store = createViewerOptionsStore();
		store.setConstraints(LOCK_OFF);
		expect(store.getOptions().trust.allowExternalContent).toBe(false);
		store.setConstraints({});
		expect(store.getOptions().trust.allowExternalContent).toBe(true);
		expect(storedTrust()).toStrictEqual({ allowExternalContent: true });
	});

	it('returns to the default when the lock is lifted and the user never chose', () => {
		const store = createViewerOptionsStore();
		store.setConstraints(LOCK_ON);
		expect(store.getOptions().trust.allowExternalContent).toBe(true);
		store.setConstraints({});
		expect(store.getOptions().trust.allowExternalContent).toBe(false);
	});

	it('restores the pre-lock value without storage', () => {
		const store = createViewerOptionsStore({ persist: false });
		store.setValue('trust', 'allowExternalContent', true);
		store.setConstraints(LOCK_OFF);
		// A lock that changes value is still the same lock: the pre-lock value is kept.
		store.setConstraints(LOCK_ON);
		store.setConstraints(LOCK_OFF);
		expect(store.getOptions().trust.allowExternalContent).toBe(false);
		store.setConstraints({});
		expect(store.getOptions().trust.allowExternalContent).toBe(true);
	});
});
