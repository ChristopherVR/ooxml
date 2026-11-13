// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { resetMetricTrackingCache } from './text-metric-tracking';
import { settleTracking, TRACKING_PENDING_VAR, trackingPendingStyle } from './text-tracking-settle';

/** Every face is `perChar` px a character: Aptos is never painted, so a stand-in always is. */
let perChar = 10;
const FAMILY = '"Aptos", "Gantari", "Calibri", sans-serif';

beforeEach(() => {
	perChar = 10;
	resetMetricTrackingCache();
	vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
		() =>
			({
				font: '',
				fontKerning: 'auto',
				measureText: (text: string) => ({ width: [...text].length * perChar }),
			}) as unknown as CanvasRenderingContext2D,
	);
});

afterEach(() => {
	vi.restoreAllMocks();
	resetMetricTrackingCache();
	document.body.innerHTML = '';
});

function markedRun(text: string, authoredPx: number, letterSpacing: string): HTMLElement {
	const span = document.createElement('span');
	span.textContent = text;
	span.style.fontFamily = FAMILY;
	span.style.fontSize = '100px';
	span.style.letterSpacing = letterSpacing;
	span.style.setProperty(TRACKING_PENDING_VAR, String(authoredPx));
	document.body.append(span);
	return span;
}

describe('trackingPendingStyle', () => {
	it('marks a run whose deck font is painted by a stand-in, with its authored spacing', () => {
		expect(trackingPendingStyle({ fontFamily: FAMILY, fontSizePx: 24 }, 1.5)).toStrictEqual({
			[TRACKING_PENDING_VAR]: '1.5',
		});
	});

	it('leaves a font without measured advances alone', () => {
		expect(trackingPendingStyle({ fontFamily: '"Papyrus", serif', fontSizePx: 24 }, 0)).toBe(
			undefined,
		);
	});
});

describe('settleTracking (issue #249)', () => {
	it('recomputes the spacing for the face that loaded after the run was built', () => {
		// Built while the narrow stand-in was painted: a wide spacing.
		const span = markedRun('dolore', 0, '20px');
		// The look-alike arrives, wider: the run needs less spacing, not the stale 20px.
		perChar = 50;
		resetMetricTrackingCache();
		expect(settleTracking()).toBe(1);
		const settled = Number.parseFloat(span.style.letterSpacing);
		expect(settled).toBeLessThan(20);
		// Still a stand-in, so a later face can settle it again.
		expect(span.style.getPropertyValue(TRACKING_PENDING_VAR)).toBe('0');
	});

	it('keeps the authored spacing on top of the measured correction', () => {
		const plain = markedRun('dolore', 0, '0px');
		const spaced = markedRun('dolore', 3, '0px');
		settleTracking();
		expect(
			Number.parseFloat(spaced.style.letterSpacing) -
				Number.parseFloat(plain.style.letterSpacing || '0'),
		).toBeCloseTo(3, 6);
	});

	it('drops a marker it cannot read', () => {
		const span = markedRun('', 0, '1px');
		settleTracking();
		expect(span.style.getPropertyValue(TRACKING_PENDING_VAR)).toBe('');
	});
});
