import { describe, it, expect } from 'vitest';

import { msToMmSs, validateTrimRange } from './MediaInspector';

// ---------------------------------------------------------------------------
// msToMmSs
// ---------------------------------------------------------------------------

describe('msToMmSs', () => {
	it('converts 0 ms to 00:00', () => {
		expect(msToMmSs(0)).toBe('00:00');
	});

	it('converts 1000 ms to 00:01', () => {
		expect(msToMmSs(1000)).toBe('00:01');
	});

	it('converts 60000 ms to 01:00', () => {
		expect(msToMmSs(60000)).toBe('01:00');
	});

	it('converts 65000 ms to 01:05', () => {
		expect(msToMmSs(65000)).toBe('01:05');
	});

	it('converts 3661000 ms to 61:01 (minutes can exceed 59)', () => {
		expect(msToMmSs(3661000)).toBe('61:01');
	});

	it('converts 90000 ms to 01:30', () => {
		expect(msToMmSs(90000)).toBe('01:30');
	});

	it('pads single-digit minutes and seconds', () => {
		expect(msToMmSs(5000)).toBe('00:05');
		expect(msToMmSs(300000)).toBe('05:00');
	});

	it('rounds to the nearest second', () => {
		// 1500 ms = 1.5 seconds -> rounds to 2
		expect(msToMmSs(1500)).toBe('00:02');
		// 1499 ms = 1.499 seconds -> rounds to 1
		expect(msToMmSs(1499)).toBe('00:01');
	});

	it('returns 00:00 for negative values', () => {
		expect(msToMmSs(-1000)).toBe('00:00');
		expect(msToMmSs(-1)).toBe('00:00');
	});

	it('returns 00:00 for NaN', () => {
		expect(msToMmSs(NaN)).toBe('00:00');
	});

	it('returns 00:00 for Infinity', () => {
		expect(msToMmSs(Infinity)).toBe('00:00');
		expect(msToMmSs(-Infinity)).toBe('00:00');
	});
});

// ---------------------------------------------------------------------------
// mmSsToMs
// ---------------------------------------------------------------------------

describe('validateTrimRange', () => {
	it('returns null for valid range where start < end < duration', () => {
		expect(validateTrimRange(0, 5000, 10000)).toBeNull();
	});

	it('returns null when both trimStart and trimEnd are 0 (no trim)', () => {
		expect(validateTrimRange(0, 0, 10000)).toBeNull();
	});

	it('returns null when trimEnd is 0 (unset) and trimStart is valid', () => {
		expect(validateTrimRange(1000, 0, 10000)).toBeNull();
	});

	it('returns error key for negative trimStart', () => {
		expect(validateTrimRange(-1, 5000, 10000)).toBe('pptx.media.trimErrorNegative');
	});

	it('returns error key for negative trimEnd', () => {
		expect(validateTrimRange(0, -1, 10000)).toBe('pptx.media.trimErrorNegative');
	});

	it('returns error key when start >= end (and end > 0)', () => {
		// trimEndMs is a distance from the tail: 5000 off a 10s clip ends at 5s.
		expect(validateTrimRange(5000, 5000, 10000)).toBe('pptx.media.trimErrorStartAfterEnd');
		expect(validateTrimRange(6000, 5000, 10000)).toBe('pptx.media.trimErrorStartAfterEnd');
		expect(validateTrimRange(4000, 5000, 10000)).toBeNull();
	});

	it('returns error key when start > duration', () => {
		expect(validateTrimRange(15000, 0, 10000)).toBe('pptx.media.trimErrorBeyondDuration');
	});

	it('returns error key when end > duration', () => {
		expect(validateTrimRange(0, 15000, 10000)).toBe('pptx.media.trimErrorBeyondDuration');
	});

	it('returns null when duration is 0 (unknown duration)', () => {
		// When duration is 0, duration-based checks are skipped
		expect(validateTrimRange(1000, 5000, 0)).toBeNull();
	});

	it('rejects a tail trim that swallows the whole clip', () => {
		// 10000 off the tail of a 10s clip leaves nothing to play.
		expect(validateTrimRange(0, 10000, 10000)).toBe('pptx.media.trimErrorStartAfterEnd');
		expect(validateTrimRange(0, 9000, 10000)).toBeNull();
	});
});
