import { describe, it, expect } from 'vitest';

import { formatBytes, isBrowserOpenableDataUrl, oleActionsVisible } from './ole-actions';

describe('formatBytes', () => {
	it('returns undefined for missing or invalid input', () => {
		expect(formatBytes(undefined)).toBeUndefined();
		expect(formatBytes(-1)).toBeUndefined();
		expect(formatBytes(Number.NaN)).toBeUndefined();
		expect(formatBytes(Number.POSITIVE_INFINITY)).toBeUndefined();
	});

	it('formats small byte counts with correct pluralisation', () => {
		expect(formatBytes(0)).toBe('0 bytes');
		expect(formatBytes(1)).toBe('1 byte');
		expect(formatBytes(512)).toBe('512 bytes');
	});

	it('formats KB/MB/GB with binary units', () => {
		expect(formatBytes(1024)).toBe('1 KB');
		expect(formatBytes(1536)).toBe('1.5 KB');
		expect(formatBytes(1024 * 1024)).toBe('1 MB');
		expect(formatBytes(12 * 1024 * 1024)).toBe('12 MB');
		expect(formatBytes(1024 * 1024 * 1024)).toBe('1 GB');
	});
});

describe('isBrowserOpenableDataUrl', () => {
	it('opens pdf, raster images, plain text and json payloads', () => {
		for (const type of [
			'application/pdf',
			'image/png',
			'text/plain',
			'application/json',
			'IMAGE/JPEG',
		]) {
			expect(isBrowserOpenableDataUrl(`data:${type};base64,AAAA`)).toBe(true);
		}
		expect(isBrowserOpenableDataUrl('data:TEXT/Plain;charset=UTF-8,hi')).toBe(true);
	});

	it('refuses markup types that can run script on the viewer origin', () => {
		for (const type of [
			'image/svg+xml',
			'text/html',
			'text/xml',
			'application/xml',
			'application/xhtml+xml',
			'application/vnd.custom+xml',
		]) {
			expect(isBrowserOpenableDataUrl(`data:${type},x`)).toBe(false);
		}
	});

	it('reads the type from the URL, not from anything a field claims', () => {
		expect(isBrowserOpenableDataUrl('javascript:alert(1)')).toBe(false);
		expect(isBrowserOpenableDataUrl('https://example.com/a.pdf')).toBe(false);
		expect(isBrowserOpenableDataUrl(undefined)).toBe(false);
		expect(
			isBrowserOpenableDataUrl(
				'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,AA',
			),
		).toBe(false);
	});
});

describe('oleActionsVisible', () => {
	it('shows the action bar only on the interactive editable canvas', () => {
		expect(oleActionsVisible({ presenting: false, preview: false })).toBeTruthy();
	});

	it('hides it on a still of a slide (thumbnail, presenter pane, export raster)', () => {
		// This is the nested-<button> bug: a slide thumbnail is rendered inside
		// a `<button aria-label="Go to slide N">` in every binding, so a
		// Download/Open `<button>` painted here is invalid, un-clickable markup.
		expect(oleActionsVisible({ presenting: false, preview: true })).toBeFalsy();
	});

	it('hides it on the live presentation stage', () => {
		expect(oleActionsVisible({ presenting: true, preview: false })).toBeFalsy();
	});
});
