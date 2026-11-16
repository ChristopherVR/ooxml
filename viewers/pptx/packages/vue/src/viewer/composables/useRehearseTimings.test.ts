// oxlint-disable react-hooks/rules-of-hooks
import { describe, expect, it, vi } from 'vitest';
import { effectScope } from 'vue';

import { formatRehearseMs, useRehearseTimings } from './useRehearseTimings';

describe('formatRehearseMs', () => {
	it('formats as M:SS, flooring and clamping', () => {
		expect(formatRehearseMs(0)).toBe('0:00');
		expect(formatRehearseMs(65000)).toBe('1:05');
		expect(formatRehearseMs(1999)).toBe('0:01');
		expect(formatRehearseMs(-500)).toBe('0:00');
	});
});

describe('useRehearseTimings (shared rehearsal controller)', () => {
	it('records each slide and saves through onSave', () => {
		let t = 0;
		const onSave = vi.fn();
		const scope = effectScope();
		const rehearsal = scope.run(() => useRehearseTimings({ now: () => t, onSave }))!;
		rehearsal.start(1);
		expect(rehearsal.rehearsing.value).toBe(true);
		t = 2000;
		rehearsal.moveTo(2);
		t = 2500;
		rehearsal.togglePause();
		t = 9000;
		rehearsal.togglePause();
		t = 9500;
		rehearsal.finish();
		expect(rehearsal.showSummary.value).toBe(true);
		expect(rehearsal.recordedTimings.value).toEqual({ 1: 2000, 2: 1000 });
		expect(rehearsal.totalRecordedMs.value).toBe(3000);
		rehearsal.saveTimings();
		expect(onSave).toHaveBeenCalledWith({ 1: 2000, 2: 1000 });
		expect(rehearsal.showSummary.value).toBe(false);
		scope.stop();
	});

	it('discards the timings', () => {
		const scope = effectScope();
		const rehearsal = scope.run(() => useRehearseTimings({ now: () => 0 }))!;
		rehearsal.start(0);
		rehearsal.finish();
		rehearsal.dismissSummary();
		expect(rehearsal.recordedTimings.value).toEqual({});
		expect(rehearsal.rehearsing.value).toBe(false);
		scope.stop();
	});
});
