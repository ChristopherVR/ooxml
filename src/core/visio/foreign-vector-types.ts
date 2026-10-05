/** Inert converter-output scene. Never contains markup, document IDs, URLs, fonts or images. */
export type VisioForeignVectorMatrix = readonly [number, number, number, number, number, number];
export type VisioForeignVectorRule = 'nonzero' | 'evenodd';
/** Commands are normalized to absolute coordinates; each values array has the command's SVG arity. */
export interface VisioForeignVectorCommand {
	readonly command: 'M' | 'L' | 'C' | 'Q' | 'A' | 'Z';
	readonly values: readonly number[];
}
export interface VisioForeignVectorPaint {
	readonly fill: string;
	readonly stroke: string;
	readonly fillRule: VisioForeignVectorRule;
	readonly strokeWidth: number;
	readonly strokeMiterlimit: number;
	readonly strokeLinecap: 'butt' | 'round' | 'square';
	readonly strokeLinejoin: 'miter' | 'round' | 'bevel';
	readonly opacity: number;
	readonly fillOpacity: number;
	readonly strokeOpacity: number;
}
export interface VisioForeignVectorPath {
	readonly kind: 'path';
	readonly matrix: VisioForeignVectorMatrix;
	readonly commands: readonly VisioForeignVectorCommand[];
	readonly paint: VisioForeignVectorPaint;
	/** Exact index into this scene's clips. No cross-scene or DOM identity. */
	readonly clipIndex?: number;
}
export interface VisioForeignVectorGroup {
	readonly kind: 'group';
	readonly matrix: VisioForeignVectorMatrix;
	readonly items: readonly VisioForeignVectorNode[];
	readonly clipIndex?: number;
}
export type VisioForeignVectorNode = VisioForeignVectorPath | VisioForeignVectorGroup;
export interface VisioForeignVectorClipPath {
	readonly matrix: VisioForeignVectorMatrix;
	readonly commands: readonly VisioForeignVectorCommand[];
	/** SVG clip-rule, independently defaulted to nonzero, never inferred from fill-rule. */
	readonly clipRule: VisioForeignVectorRule;
	readonly clipIndex?: number;
}
export interface VisioForeignVectorClip {
	readonly matrix: VisioForeignVectorMatrix;
	/** Union of path silhouettes. An empty array clips away everything.
	 * https://www.w3.org/TR/SVG11/masking.html#EstablishingANewClippingPath
	 */
	readonly items: readonly VisioForeignVectorClipPath[];
	/** Intersection with another validated clip, in userSpaceOnUse coordinates. */
	readonly clipIndex?: number;
}
export interface VisioForeignVector {
	readonly kind: 'vector';
	readonly width: number;
	readonly height: number;
	readonly items: readonly VisioForeignVectorNode[];
	readonly clips: readonly VisioForeignVectorClip[];
}
export interface VisioForeignVectorLimits {
	maxNodes: number;
	/** Aggregate UTF-16 code units of source SVG strings and, separately, retained scene values. */
	maxCharacters: number;
	maxPathOperands: number;
	maxPathCommands: number;
	maxDepth: number;
	maxDimension: number;
	maxCoordinate: number;
	/** Includes every clip reference expansion, even repeated uses. */
	maxExpandedNodes: number;
	maxExpandedOperands: number;
	maxExpandedCommands: number;
}
export const VISIO_FOREIGN_VECTOR_LIMITS: Readonly<VisioForeignVectorLimits> = Object.freeze({
	maxNodes: 25_000,
	maxCharacters: 2 * 1024 * 1024,
	maxPathOperands: 100_000,
	maxPathCommands: 50_000,
	maxDepth: 64,
	maxDimension: 2048,
	maxCoordinate: 1_000_000,
	maxExpandedNodes: 25_000,
	maxExpandedOperands: 100_000,
	maxExpandedCommands: 50_000,
});
export class VisioForeignVectorError extends Error {
	override readonly name = 'VisioForeignVectorError';
	constructor(
		readonly code: 'unsafe-vector' | 'vector-limit',
		message: string,
	) {
		super(message);
	}
}
export function unsafe(message: string): never {
	throw new VisioForeignVectorError('unsafe-vector', message);
}
export function limit(message: string): never {
	throw new VisioForeignVectorError('vector-limit', message);
}
