declare const unitBrand: unique symbol;
declare const nonNegativeBrand: unique symbol;
type Branded<Name extends string> = number & {
    readonly [unitBrand]: Name;
};
/** Twentieths of a point (1/1440 inch), possibly negative: `ST_SignedTwipsMeasure`. */
export type SignedTwips = Branded<'twips'>;
/**
 * Non-negative twentieths of a point (`ST_TwipsMeasure`): WordprocessingML's native length unit.
 * A `Twips` is assignable to `SignedTwips`, never the other way round.
 */
export type Twips = SignedTwips & {
    readonly [nonNegativeBrand]: true;
};
/** Half-points (`w:sz`): font sizes are stored as twice their point size. */
export type HalfPoints = Branded<'halfPoints'>;
/** Eighth-points (`w:sz` on borders): border widths. */
export type EighthPoints = Branded<'eighthPoints'>;
/** English Metric Units (1/914400 inch): DrawingML lengths. */
export type Emu = Branded<'emu'>;
/** A non-negative whole number of twips; throws `RangeError` otherwise. */
export declare const twips: (value: number) => Twips;
/** A whole number of twips that may be negative; throws `RangeError` unless a safe integer. */
export declare const signedTwips: (value: number) => SignedTwips;
export declare const halfPoints: (value: number) => HalfPoints;
export declare const eighthPoints: (value: number) => EighthPoints;
export declare const emu: (value: number) => Emu;
export declare const roundTwips: (value: number) => Twips;
export declare const roundSignedTwips: (value: number) => SignedTwips;
export declare const roundHalfPoints: (value: number) => HalfPoints;
export declare const roundEighthPoints: (value: number) => EighthPoints;
export declare const roundEmu: (value: number) => Emu;
/** Points (1/72 inch) to twips; negative input clamps to 0. */
export declare const twipsFromPoints: (points: number) => Twips;
export declare const signedTwipsFromPoints: (points: number) => SignedTwips;
/** CSS pixels at 96 dpi (15 twips per pixel) to twips; negative input clamps to 0. */
export declare const twipsFromPixels: (pixels: number) => Twips;
export declare const signedTwipsFromPixels: (pixels: number) => SignedTwips;
export declare const twipsFromInches: (inches: number) => Twips;
/** A font size in points to half-points (Word stores `w:sz` as twice the point size). */
export declare const halfPointsFromPoints: (points: number) => HalfPoints;
/** A border width in points to eighth-points; negative input clamps to 0. */
export declare const eighthPointsFromPoints: (points: number) => EighthPoints;
export declare const emuFromPixels: (pixels: number) => Emu;
export declare const twipsToPoints: (value: SignedTwips) => number;
export declare const twipsToPixels: (value: SignedTwips) => number;
export declare const twipsToInches: (value: SignedTwips) => number;
export declare const halfPointsToPoints: (value: HalfPoints) => number;
export declare const eighthPointsToPoints: (value: EighthPoints) => number;
export declare const emuToPixels: (value: Emu) => number;
export {};
