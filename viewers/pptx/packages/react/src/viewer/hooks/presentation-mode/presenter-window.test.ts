import { describe, it, expect, expectTypeOf } from 'vitest';

import {
	isPresenterMessage,
	PRESENTER_MSG_ORIGIN,
	PRESENTER_CHANNEL_NAME,
	isAudienceTab,
	AUDIENCE_HASH,
} from './usePresenterWindow';
import type {
	PresenterSlideChangeMessage,
	PresenterExitMessage,
	PresenterMessage,
} from './usePresenterWindow';

// ---------------------------------------------------------------------------
// The window and channel logic is the shared presenter session controller,
// tested in ooxml-ui/pptx; these pin the exported React helpers.
// ---------------------------------------------------------------------------

const TEST_SESSION_ID = '00000000-0000-0000-0000-000000000001';

// ---------------------------------------------------------------------------
// isPresenterMessage
// ---------------------------------------------------------------------------

describe('isPresenterMessage', () => {
	it('accepts a valid slide-change message', () => {
		const msg: PresenterSlideChangeMessage = {
			origin: PRESENTER_MSG_ORIGIN,
			type: 'presenter-slide-change',
			slideIndex: 3,
			sessionId: TEST_SESSION_ID,
		};
		expect(isPresenterMessage(msg)).toBeTruthy();
	});

	it('accepts a valid exit message', () => {
		const msg: PresenterExitMessage = {
			origin: PRESENTER_MSG_ORIGIN,
			type: 'presenter-exit',
			sessionId: TEST_SESSION_ID,
		};
		expect(isPresenterMessage(msg)).toBeTruthy();
	});

	it('rejects null', () => {
		expect(isPresenterMessage(null)).toBeFalsy();
	});

	it('rejects undefined', () => {
		expect(isPresenterMessage(undefined)).toBeFalsy();
	});

	it('rejects a string', () => {
		expect(isPresenterMessage('hello')).toBeFalsy();
	});

	it('rejects a number', () => {
		expect(isPresenterMessage(42)).toBeFalsy();
	});

	it('rejects an object with wrong origin', () => {
		expect(
			isPresenterMessage({
				origin: 'wrong-origin',
				type: 'presenter-slide-change',
				slideIndex: 0,
				sessionId: TEST_SESSION_ID,
			}),
		).toBeFalsy();
	});

	it('rejects an object with unknown type', () => {
		expect(
			isPresenterMessage({
				origin: PRESENTER_MSG_ORIGIN,
				type: 'unknown-type',
				sessionId: TEST_SESSION_ID,
			}),
		).toBeFalsy();
	});

	it('rejects a message missing the sessionId', () => {
		expect(
			isPresenterMessage({
				origin: PRESENTER_MSG_ORIGIN,
				type: 'presenter-slide-change',
				slideIndex: 0,
			}),
		).toBeFalsy();
	});

	it('rejects an empty object', () => {
		expect(isPresenterMessage({})).toBeFalsy();
	});

	it('rejects an object with only the origin', () => {
		expect(isPresenterMessage({ origin: PRESENTER_MSG_ORIGIN })).toBeFalsy();
	});
});

// ---------------------------------------------------------------------------
// isAudienceTab: requires a browser environment
// ---------------------------------------------------------------------------

describe('isAudienceTab', () => {
	it('is exported as a function', () => {
		expectTypeOf(isAudienceTab).toBeFunction();
	});

	it('aUDIENCE_HASH is the expected value', () => {
		expect(AUDIENCE_HASH).toBe('#pptx-audience');
	});
});
