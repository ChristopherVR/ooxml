import type { DrawingColor, DrawingFill } from './types';

interface EffectSource {
	sourceXml?: string;
}

/** Shadow geometry stays in OOXML units until a painter resolves it. */
export interface DrawingShadowGeometry {
	blurRadiusEmu?: number;
	distanceEmu?: number;
	/** Angle in 1/60000 degrees. */
	directionAngle?: number;
	color?: DrawingColor;
}

export interface DrawingOuterShadow extends EffectSource, DrawingShadowGeometry {
	kind: 'outerShadow';
	/** Scale in 1/1000 percent. */
	scaleX?: number;
	scaleY?: number;
	skewXAngle?: number;
	skewYAngle?: number;
	alignment?: string;
	rotateWithShape?: boolean;
}

export interface DrawingReflection extends EffectSource {
	kind: 'reflection';
	blurRadiusEmu?: number;
	distanceEmu?: number;
	directionAngle?: number;
	fadeDirectionAngle?: number;
	startAlpha?: number;
	startPosition?: number;
	endAlpha?: number;
	endPosition?: number;
	scaleX?: number;
	scaleY?: number;
	skewXAngle?: number;
	skewYAngle?: number;
	alignment?: string;
	rotateWithShape?: boolean;
}

/** CT_EffectList choices, in authored order. DAGs and unknown effects remain opaque. */
export type DrawingEffect =
	| DrawingOuterShadow
	| DrawingReflection
	| (EffectSource & DrawingShadowGeometry & { kind: 'innerShadow' })
	| (EffectSource &
			Omit<DrawingShadowGeometry, 'blurRadiusEmu'> & { kind: 'presetShadow'; preset?: string })
	| (EffectSource & { kind: 'glow'; radiusEmu?: number; color?: DrawingColor })
	| (EffectSource & { kind: 'softEdge'; radiusEmu?: number })
	| (EffectSource & { kind: 'blur'; radiusEmu?: number; grow?: boolean })
	| (EffectSource & { kind: 'fillOverlay'; blend?: string; fill?: DrawingFill })
	| { kind: 'unsupported'; sourceXml: string };

export interface DrawingEffectList extends EffectSource {
	effects: DrawingEffect[];
}
