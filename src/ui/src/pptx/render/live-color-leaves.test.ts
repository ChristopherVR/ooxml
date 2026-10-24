import { describe, expect, it } from 'vitest';

import { liveColorRevertPatch } from './live-color-leaves';

describe('liveColorRevertPatch', () => {
	it('puts back only the leaves named, on the value as it is now', () => {
		const now = { style: { fill: '#330000', width: 5, glow: 'x' }, text: 'Hi' };
		expect(
			liveColorRevertPatch(now, [
				{ path: ['style', 'fill'], present: true, value: '#4472c4' },
				{ path: ['style', 'glow'], present: false, value: undefined },
			]),
		).toStrictEqual({ style: { fill: '#4472c4', width: 5 } });
	});

	it('brings back a top-level entry the target did not have as undefined', () => {
		expect(
			liveColorRevertPatch({ style: { fill: '#000' } }, [
				{ path: ['style'], present: false, value: undefined },
			]),
		).toStrictEqual({ style: undefined });
		expect(
			liveColorRevertPatch(undefined, [{ path: ['a', 'b'], present: true, value: 1 }]),
		).toStrictEqual({ a: { b: 1 } });
	});

	it('writes into arrays by index', () => {
		const now = { runs: [{ color: '#330000' }, { color: '#330000', bold: true }] };
		expect(
			liveColorRevertPatch(now, [
				{ path: ['runs', '1', 'color'], present: true, value: '#112233' },
			]),
		).toStrictEqual({ runs: [{ color: '#330000' }, { color: '#112233', bold: true }] });
	});
});
