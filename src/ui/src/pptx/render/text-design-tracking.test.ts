import { afterEach, describe, expect, it } from 'vitest';

import {
	deckFamilyOf,
	resetDesignTrackingCache,
	resolveDesignTrackingPx,
} from './text-design-tracking';

/** A canvas whose text is 10px per character in a stand-in and 12px per character in `Aptos`. */
function stubContext(aptosInstalled: boolean): CanvasRenderingContext2D {
	const ctx = { font: '' } as { font: string; measureText: (text: string) => { width: number } };
	ctx.measureText = (text) => ({
		width: text.length * (aptosInstalled && ctx.font.includes('"Aptos"') ? 12 : 10),
	});
	return ctx as unknown as CanvasRenderingContext2D;
}

afterEach(resetDesignTrackingCache);

describe('deckFamilyOf', () => {
	it('takes the first family of a CSS list, unquoted', () => {
		expect(deckFamilyOf('"Aptos", "Calibri", sans-serif')).toBe('Aptos');
		expect(deckFamilyOf('Aptos, Calibri')).toBe('Aptos');
		expect(deckFamilyOf(undefined)).toBeUndefined();
	});
});

describe('resolveDesignTrackingPx', () => {
	it("spreads Aptos's measured width over a stand-in that is narrower", () => {
		const tracking = resolveDesignTrackingPx(stubContext(false), 'Hi', 'Aptos', 100, false, 20);
		expect(tracking).toBeGreaterThan(0);
	});

	it('leaves a deck font the browser paints to the rounding model', () => {
		expect(
			resolveDesignTrackingPx(stubContext(true), 'Hi', 'Aptos', 100, false, 20),
		).toBeUndefined();
	});

	it('leaves a font that was not measured, and text outside the measured range, alone', () => {
		expect(
			resolveDesignTrackingPx(stubContext(false), 'Hi', 'Papyrus', 100, false, 20),
		).toBeUndefined();
		expect(
			resolveDesignTrackingPx(stubContext(false), 'Hé', 'Aptos', 100, false, 20),
		).toBeUndefined();
	});

	it('uses the bold advances for bold text', () => {
		const regular = resolveDesignTrackingPx(stubContext(false), 'Hello', 'Aptos', 100, false, 250)!;
		const bold = resolveDesignTrackingPx(stubContext(false), 'Hello', 'Aptos', 100, true, 250)!;
		expect(bold).toBeGreaterThan(regular);
	});
});
