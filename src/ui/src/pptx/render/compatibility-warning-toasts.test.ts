import { describe, expect, it } from 'vitest';

import {
	compatibilityWarningToasts,
	createRemoteValueNotifier,
} from './compatibility-warning-toasts';

describe('compatibilityWarningToasts', () => {
	it('maps a known code to its message key', () => {
		const toasts = compatibilityWarningToasts([
			{ code: 'SAVE_SIGNATURES_STRIPPED', severity: 'warning' },
		]);
		expect(toasts).toStrictEqual([
			{
				id: 'SAVE_SIGNATURES_STRIPPED',
				code: 'SAVE_SIGNATURES_STRIPPED',
				severity: 'warning',
				messageKey: 'pptx.compatibility.saveSignaturesStripped',
				params: undefined,
			},
		]);
	});

	it('maps EXTERNAL_IMAGE_REFERENCE', () => {
		const toasts = compatibilityWarningToasts([
			{ code: 'EXTERNAL_IMAGE_REFERENCE', severity: 'info' },
		]);
		expect(toasts[0].messageKey).toBe('pptx.compatibility.externalImageReference');
		expect(toasts[0].severity).toBe('info');
	});

	it('maps the embedded font licence codes', () => {
		const toasts = compatibilityWarningToasts([
			{ code: 'EMBEDDED_FONT_RESTRICTED', severity: 'warning' },
			{ code: 'EMBEDDED_FONT_NOT_EMBEDDABLE', severity: 'warning' },
		]);
		expect(toasts.map((toast) => toast.messageKey)).toStrictEqual([
			'pptx.compatibility.embeddedFontRestricted',
			'pptx.compatibility.embeddedFontNotEmbeddable',
		]);
	});

	// The graphic-frame limitation codes used to fall through to the generic message,
	// which shows the raw code to the user.
	it.each([
		['PARTIAL_OLE_SUPPORT', 'pptx.compatibility.partialOleSupport'],
		['PARTIAL_INK_SUPPORT', 'pptx.compatibility.partialInkSupport'],
		['UNSUPPORTED_GRAPHIC_FRAME', 'pptx.compatibility.unsupportedGraphicFrame'],
	])('maps %s to its own message', (code, messageKey) => {
		const [toast] = compatibilityWarningToasts([{ code, severity: 'info' }]);
		expect(toast.messageKey).toBe(messageKey);
		expect(toast.params).toBeUndefined();
	});

	it('dedupes identical codes, keeping the first severity', () => {
		const toasts = compatibilityWarningToasts([
			{ code: 'SAVE_ELEMENT_SKIPPED', severity: 'warning', elementId: 'a' },
			{ code: 'SAVE_ELEMENT_SKIPPED', severity: 'info', elementId: 'b' },
			{ code: 'SAVE_ELEMENT_SKIPPED', severity: 'warning', elementId: 'c' },
		]);
		expect(toasts).toHaveLength(1);
		expect(toasts[0].severity).toBe('warning');
	});

	it('falls back to the generic key with a code param for an unknown code', () => {
		const toasts = compatibilityWarningToasts([{ code: 'SOME_FUTURE_CODE', severity: 'warning' }]);
		expect(toasts[0].messageKey).toBe('pptx.compatibility.generic');
		expect(toasts[0].params).toStrictEqual({ code: 'SOME_FUTURE_CODE' });
	});

	it('returns one toast per distinct code, preserving first-seen order', () => {
		const toasts = compatibilityWarningToasts([
			{ code: 'UNMODELLED_SLIDE_MARKUP', severity: 'info' },
			{ code: 'EXTERNAL_IMAGE_REFERENCE', severity: 'warning' },
			{ code: 'UNMODELLED_SLIDE_MARKUP', severity: 'info' },
		]);
		expect(toasts.map((t) => t.code)).toStrictEqual([
			'UNMODELLED_SLIDE_MARKUP',
			'EXTERNAL_IMAGE_REFERENCE',
		]);
	});

	it('returns an empty list for no warnings', () => {
		expect(compatibilityWarningToasts([])).toStrictEqual([]);
	});

	it('maps the lowercase-hyphen group-depth-exceeded code', () => {
		const toasts = compatibilityWarningToasts([
			{ code: 'group-depth-exceeded', severity: 'warning' },
		]);
		expect(toasts[0].messageKey).toBe('pptx.compatibility.groupDepthExceeded');
	});
});

describe('createRemoteValueNotifier', () => {
	const rejected = (elementId: string) => ({
		code: 'COLLABORATION_REMOTE_VALUE_REJECTED',
		severity: 'warning' as const,
		message: 'left out',
		elementId,
	});

	it('gives a translated toast once per rejection a peer caused', () => {
		const notice = createRemoteValueNotifier();
		expect(notice([{ id: 's1', warnings: [] }])).toBeUndefined();
		const toast = notice([{ id: 's1', warnings: [rejected('a')] }]);
		expect(toast?.messageKey).toBe('pptx.compatibility.collaborationValueRejected');
		expect(toast?.severity).toBe('warning');
		// The same rejection, read again with the next remote change, is not reported again.
		expect(notice([{ id: 's1', warnings: [rejected('a')] }])).toBeUndefined();
		expect(notice([{ id: 's1', warnings: [rejected('a'), rejected('b')] }])?.id).toBe(
			'COLLABORATION_REMOTE_VALUE_REJECTED',
		);
	});
});
