/**
 * 3D-text (WordArt format-text-effects) inspector option lists: the OOXML bevel
 * preset catalogue and the surface-material preset catalogue, each with display
 * labels. Framework-free pure data, consumed by every binding's Text 3D
 * properties panel (React / Vue / Angular).
 */
import type { BevelPresetType, MaterialPresetType } from 'ooxml-core/pptx';

export const BEVEL_PRESETS: ReadonlyArray<{
	value: BevelPresetType;
	label: string;
	/** Translation key for `label`; bindings paint the translated name. */
	labelKey: string;
}> = [
	{ value: 'none', label: 'None', labelKey: 'pptx.bevelTypes.none' },
	{ value: 'circle', label: 'Circle', labelKey: 'pptx.bevelTypes.circle' },
	{ value: 'relaxedInset', label: 'Relaxed Inset', labelKey: 'pptx.bevelTypes.relaxedInset' },
	{ value: 'cross', label: 'Cross', labelKey: 'pptx.bevelTypes.cross' },
	{ value: 'coolSlant', label: 'Cool Slant', labelKey: 'pptx.bevelTypes.coolSlant' },
	{ value: 'angle', label: 'Angle', labelKey: 'pptx.bevelTypes.angle' },
	{ value: 'softRound', label: 'Soft Round', labelKey: 'pptx.bevelTypes.softRound' },
	{ value: 'convex', label: 'Convex', labelKey: 'pptx.bevelTypes.convex' },
	{ value: 'slope', label: 'Slope', labelKey: 'pptx.bevelTypes.slope' },
	{ value: 'divot', label: 'Divot', labelKey: 'pptx.bevelTypes.divot' },
	{ value: 'riblet', label: 'Riblet', labelKey: 'pptx.bevelTypes.riblet' },
	{ value: 'hardEdge', label: 'Hard Edge', labelKey: 'pptx.bevelTypes.hardEdge' },
	{ value: 'artDeco', label: 'Art Deco', labelKey: 'pptx.bevelTypes.artDeco' },
];

export const MATERIAL_PRESETS: ReadonlyArray<{
	value: MaterialPresetType | '';
	label: string;
	/** Translation key for `label`; bindings paint the translated name. */
	labelKey: string;
}> = [
	{ value: '', label: 'None', labelKey: 'pptx.material.none' },
	{ value: 'matte', label: 'Matte', labelKey: 'pptx.material.matte' },
	{ value: 'warmMatte', label: 'Warm Matte', labelKey: 'pptx.material.warmMatte' },
	{ value: 'plastic', label: 'Plastic', labelKey: 'pptx.material.plastic' },
	{ value: 'metal', label: 'Metal', labelKey: 'pptx.material.metal' },
	{ value: 'dkEdge', label: 'Dark Edge', labelKey: 'pptx.material.dkEdge' },
	{ value: 'softEdge', label: 'Soft Edge', labelKey: 'pptx.material.softEdge' },
	{ value: 'flat', label: 'Flat', labelKey: 'pptx.material.flat' },
	{ value: 'softmetal', label: 'Soft Metal', labelKey: 'pptx.material.softmetal' },
	{ value: 'clear', label: 'Clear', labelKey: 'pptx.material.clear' },
	{ value: 'powder', label: 'Powder', labelKey: 'pptx.material.powder' },
	{
		value: 'translucentPowder',
		label: 'Translucent Powder',
		labelKey: 'pptx.material.translucentPowder',
	},
];
