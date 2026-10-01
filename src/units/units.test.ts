import { describe, expect, expectTypeOf, it } from 'vitest';
import {
	eighthPoints,
	eighthPointsFromPoints,
	emu,
	emuToPixels,
	emuFromPixels,
	halfPoints,
	halfPointsFromPoints,
	roundEighthPoints,
	roundHalfPoints,
	roundSignedTwips,
	roundTwips,
	signedTwips,
	signedTwipsFromPixels,
	signedTwipsFromPoints,
	twips,
	twipsFromInches,
	twipsFromPixels,
	twipsFromPoints,
	twipsToInches,
	twipsToPixels,
	twipsToPoints,
	type EighthPoints,
	type Emu,
	type HalfPoints,
	type SignedTwips,
	type Twips,
	EMU_PER_INCH,
	EMU_PER_PIXEL,
	EMU_PER_POINT,
	TWIPS_PER_INCH,
	TWIPS_PER_POINT,
} from './index.js';

describe('strict constructors', () => {
	it('twips rejects negatives and fractions; signedTwips allows negatives only', () => {
		expect(twips(0)).toBe(0);
		expect(() => twips(-1)).toThrow(RangeError);
		expect(() => twips(1.5)).toThrow(RangeError);
		expect(signedTwips(-720)).toBe(-720);
		expect(() => signedTwips(0.5)).toThrow(RangeError);
		expect(() => signedTwips(Number.POSITIVE_INFINITY)).toThrow(RangeError);
		expect(() => eighthPoints(0.5)).toThrow(RangeError);
		expect(() => emu(Number.NaN)).toThrow(RangeError);
	});
});

describe('rounding constructors', () => {
	it('round to whole units, halves away from zero, without -0', () => {
		expect(roundTwips(10.4)).toBe(10);
		expect(roundTwips(10.5)).toBe(11);
		expect(roundSignedTwips(-10.5)).toBe(-11);
		expect(Object.is(roundSignedTwips(-0.2), 0)).toBe(true);
		expect(Object.is(roundTwips(-0.2), 0)).toBe(true);
		expect(roundHalfPoints(21.5)).toBe(22);
		expect(roundEighthPoints(3.6)).toBe(4);
	});
	it('roundTwips clamps negatives to zero while the signed variant keeps them', () => {
		expect(roundTwips(-40)).toBe(0);
		expect(roundSignedTwips(-40.4)).toBe(-40);
		expect(roundEighthPoints(-2)).toBe(0);
	});
	it('reject non-finite input instead of writing NaN into the model', () => {
		for (const bad of [Number.NaN, Infinity, -Infinity]) {
			expect(() => roundTwips(bad)).toThrow(RangeError);
			expect(() => roundSignedTwips(bad)).toThrow(RangeError);
			expect(() => roundHalfPoints(bad)).toThrow(RangeError);
		}
	});
});

describe('unit conversions', () => {
	it('convert from points, pixels and inches with rounding', () => {
		expect(twipsFromPoints(36)).toBe(720);
		expect(twipsFromPoints(0.26)).toBe(5);
		expect(signedTwipsFromPoints(-36)).toBe(-720);
		expect(twipsFromPixels(96)).toBe(1440);
		expect(twipsFromPixels(10.03)).toBe(150);
		expect(signedTwipsFromPixels(-7.5)).toBe(-113);
		expect(twipsFromInches(1.5)).toBe(2160);
		expect(halfPointsFromPoints(10.5)).toBe(21);
		expect(halfPointsFromPoints(10.26)).toBe(21);
		expect(eighthPointsFromPoints(0.5)).toBe(4);
		expect(emuFromPixels(96)).toBe(914400);
	});
	it('convert back to plain numbers', () => {
		expect(twipsToPoints(twips(720))).toBe(36);
		expect(twipsToPixels(signedTwips(-720))).toBe(-48);
		expect(twipsToInches(twips(2160))).toBe(1.5);
	});
	it('round-trips whole pixels through twips', () => {
		for (const px of [0, 1, 37, 816, 1056]) expect(twipsToPixels(twipsFromPixels(px))).toBe(px);
	});
});

describe('compile-time unit safety', () => {
	it('brands are distinct types that stay numbers at runtime', () => {
		expectTypeOf<Twips>().toMatchTypeOf<number>();
		expectTypeOf<Twips>().toMatchTypeOf<SignedTwips>();
		expectTypeOf<SignedTwips>().not.toMatchTypeOf<Twips>();
		expectTypeOf<HalfPoints>().not.toMatchTypeOf<Twips>();
		expectTypeOf<Twips>().not.toMatchTypeOf<HalfPoints>();
		expectTypeOf<EighthPoints>().not.toMatchTypeOf<HalfPoints>();
		expectTypeOf<Emu>().not.toMatchTypeOf<Twips>();
		expectTypeOf<number>().not.toMatchTypeOf<Twips>();
	});
	it('rejects plain numbers and other units where a brand is required', () => {
		const spacing = (value: Twips) => value;
		// @ts-expect-error a plain number is not a Twips
		spacing(240);
		// @ts-expect-error half-points are not twips
		spacing(halfPoints(24));
		// @ts-expect-error eighth-points are not twips
		spacing(eighthPoints(4));
		// @ts-expect-error a signed measure cannot go into an unsigned field
		spacing(signedTwips(-360));
		// an unsigned measure is a valid signed one
		const signed: SignedTwips = twips(720);
		expect(signed).toBe(720);
	});
	it('computed values must go through a rounding constructor', () => {
		const spacing = (value: Twips) => value;
		const dragged = 100.4;
		// @ts-expect-error an unrounded computed number cannot be used as a Twips
		spacing(dragged);
		expect(spacing(roundTwips(dragged))).toBe(100);
	});
});

describe('constants', () => {
	it('agree with the conversion helpers', () => {
		expect(EMU_PER_PIXEL).toBe(9525);
		expect(EMU_PER_INCH).toBe(EMU_PER_PIXEL * 96);
		expect(EMU_PER_POINT * 72).toBe(EMU_PER_INCH);
		expect(TWIPS_PER_INCH).toBe(TWIPS_PER_POINT * 72);
		expect(emuToPixels(emu(EMU_PER_PIXEL))).toBe(1);
	});
});
