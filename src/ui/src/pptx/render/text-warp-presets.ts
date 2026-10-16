/**
 * Text-warp (WordArt "Transform") inspector-gallery data: the full ordered list
 * of OOXML warp presets with display labels, plus a miniature SVG preview-path
 * generator used to draw each swatch. Framework-free pure data / geometry,
 * consumed by every binding's text-warp gallery (React / Vue / Angular).
 *
 * This is the UI catalogue only; the runtime warp path/CSS-transform builders
 * live in `text-warp.ts`.
 */
import type { PptxTextWarpPreset } from 'ooxml-core/pptx';

/** All standard OOXML text warp presets with display labels. */
export const TEXT_WARP_PRESETS: ReadonlyArray<{
	value: PptxTextWarpPreset;
	label: string;
	/** Translation key for `label`; bindings paint the translated name. */
	labelKey: string;
}> = [
	{ value: 'textNoShape', label: 'No Transform', labelKey: 'pptx.textWarp.preset.noShape' },
	{ value: 'textPlain', label: 'Plain', labelKey: 'pptx.textWarp.preset.plain' },
	{ value: 'textArchUp', label: 'Arch Up', labelKey: 'pptx.textWarp.preset.archUp' },
	{ value: 'textArchDown', label: 'Arch Down', labelKey: 'pptx.textWarp.preset.archDown' },
	{ value: 'textCircle', label: 'Circle', labelKey: 'pptx.textWarp.preset.circle' },
	{ value: 'textButton', label: 'Button', labelKey: 'pptx.textWarp.preset.button' },
	{ value: 'textCurveUp', label: 'Curve Up', labelKey: 'pptx.textWarp.preset.curveUp' },
	{ value: 'textCurveDown', label: 'Curve Down', labelKey: 'pptx.textWarp.preset.curveDown' },
	{ value: 'textCanUp', label: 'Can Up', labelKey: 'pptx.textWarp.preset.canUp' },
	{ value: 'textCanDown', label: 'Can Down', labelKey: 'pptx.textWarp.preset.canDown' },
	{ value: 'textWave1', label: 'Wave 1', labelKey: 'pptx.textWarp.preset.wave1' },
	{ value: 'textWave2', label: 'Wave 2', labelKey: 'pptx.textWarp.preset.wave2' },
	{ value: 'textWave4', label: 'Wave 4', labelKey: 'pptx.textWarp.preset.wave4' },
	{ value: 'textDoubleWave1', label: 'Double Wave', labelKey: 'pptx.textWarp.preset.doubleWave1' },
	{ value: 'textInflate', label: 'Inflate', labelKey: 'pptx.textWarp.preset.inflate' },
	{ value: 'textDeflate', label: 'Deflate', labelKey: 'pptx.textWarp.preset.deflate' },
	{
		value: 'textInflateBottom',
		label: 'Inflate Bottom',
		labelKey: 'pptx.textWarp.preset.inflateBottom',
	},
	{ value: 'textInflateTop', label: 'Inflate Top', labelKey: 'pptx.textWarp.preset.inflateTop' },
	{
		value: 'textDeflateBottom',
		label: 'Deflate Bottom',
		labelKey: 'pptx.textWarp.preset.deflateBottom',
	},
	{ value: 'textDeflateTop', label: 'Deflate Top', labelKey: 'pptx.textWarp.preset.deflateTop' },
	{ value: 'textFadeUp', label: 'Fade Up', labelKey: 'pptx.textWarp.preset.fadeUp' },
	{ value: 'textFadeDown', label: 'Fade Down', labelKey: 'pptx.textWarp.preset.fadeDown' },
	{ value: 'textFadeLeft', label: 'Fade Left', labelKey: 'pptx.textWarp.preset.fadeLeft' },
	{ value: 'textFadeRight', label: 'Fade Right', labelKey: 'pptx.textWarp.preset.fadeRight' },
	{ value: 'textSlantUp', label: 'Slant Up', labelKey: 'pptx.textWarp.preset.slantUp' },
	{ value: 'textSlantDown', label: 'Slant Down', labelKey: 'pptx.textWarp.preset.slantDown' },
	{ value: 'textCascadeUp', label: 'Cascade Up', labelKey: 'pptx.textWarp.preset.cascadeUp' },
	{ value: 'textCascadeDown', label: 'Cascade Down', labelKey: 'pptx.textWarp.preset.cascadeDown' },
	{ value: 'textTriangle', label: 'Triangle', labelKey: 'pptx.textWarp.preset.triangle' },
	{
		value: 'textTriangleInverted',
		label: 'Triangle Inverted',
		labelKey: 'pptx.textWarp.preset.triangleInverted',
	},
	{ value: 'textChevron', label: 'Chevron', labelKey: 'pptx.textWarp.preset.chevron' },
	{
		value: 'textChevronInverted',
		label: 'Chevron Inverted',
		labelKey: 'pptx.textWarp.preset.chevronInverted',
	},
	{ value: 'textRingInside', label: 'Ring Inside', labelKey: 'pptx.textWarp.preset.ringInside' },
	{ value: 'textRingOutside', label: 'Ring Outside', labelKey: 'pptx.textWarp.preset.ringOutside' },
	{ value: 'textStop', label: 'Stop', labelKey: 'pptx.textWarp.preset.stop' },
	{ value: 'textArchUpPour', label: 'Arch Up Pour', labelKey: 'pptx.textWarp.preset.archUpPour' },
	{
		value: 'textArchDownPour',
		label: 'Arch Down Pour',
		labelKey: 'pptx.textWarp.preset.archDownPour',
	},
	{ value: 'textCirclePour', label: 'Circle Pour', labelKey: 'pptx.textWarp.preset.circlePour' },
	{ value: 'textButtonPour', label: 'Button Pour', labelKey: 'pptx.textWarp.preset.buttonPour' },
];

/** Miniature SVG preview path for a warp preset (simplified visual hint). */
export function warpPreviewPath(preset: PptxTextWarpPreset): string {
	const w = 40;
	const h = 20;
	switch (preset) {
		case 'textNoShape':
		case 'textPlain':
			return `M 2,${h / 2} L ${w - 2},${h / 2}`;
		case 'textArchUp':
		case 'textArchUpPour':
			return `M 2,${h - 2} A ${w / 2 - 2},${h - 4} 0 0,1 ${w - 2},${h - 2}`;
		case 'textArchDown':
		case 'textArchDownPour':
			return `M 2,2 A ${w / 2 - 2},${h - 4} 0 0,0 ${w - 2},2`;
		case 'textCircle':
		case 'textCirclePour':
			return `M ${w / 2},2 A ${w / 2 - 2},${h / 2 - 2} 0 1,1 ${w / 2},${h - 2} A ${w / 2 - 2},${h / 2 - 2} 0 1,1 ${w / 2},2`;
		case 'textButton':
		case 'textButtonPour':
			return `M 2,${h / 2} Q ${w / 2},2 ${w - 2},${h / 2}`;
		case 'textCurveUp':
			return `M 2,${h - 4} Q ${w / 2},2 ${w - 2},${h - 4}`;
		case 'textCurveDown':
			return `M 2,4 Q ${w / 2},${h - 2} ${w - 2},4`;
		case 'textWave1':
			return `M 2,${h / 2} C ${w / 3},${h / 2 - 6} ${(2 * w) / 3},${h / 2 + 6} ${w - 2},${h / 2}`;
		case 'textWave2':
			return `M 2,${h / 2} C ${w / 3},${h / 2 + 6} ${(2 * w) / 3},${h / 2 - 6} ${w - 2},${h / 2}`;
		case 'textWave4':
		case 'textDoubleWave1':
			return `M 2,${h / 2} C ${w / 4},${h / 2 - 5} ${w / 2},${h / 2 + 5} ${w / 2},${h / 2} C ${w / 2},${h / 2 - 5} ${(3 * w) / 4},${h / 2 + 5} ${w - 2},${h / 2}`;
		case 'textInflate':
		case 'textInflateBottom':
		case 'textInflateTop':
			return `M 2,${h / 2} Q ${w / 2},2 ${w - 2},${h / 2}`;
		case 'textDeflate':
		case 'textDeflateBottom':
		case 'textDeflateTop':
			return `M 2,${h / 2} Q ${w / 2},${h - 2} ${w - 2},${h / 2}`;
		case 'textSlantUp':
			return `M 2,${h - 4} L ${w - 2},4`;
		case 'textSlantDown':
			return `M 2,4 L ${w - 2},${h - 4}`;
		case 'textCascadeUp':
			return `M 2,${h - 4} L ${w - 2},4`;
		case 'textCascadeDown':
			return `M 2,4 L ${w - 2},${h - 4}`;
		case 'textFadeUp':
			return `M ${w / 4},${h - 2} L ${w / 2},2 L ${(3 * w) / 4},${h - 2}`;
		case 'textFadeDown':
			return `M ${w / 4},2 L ${w / 2},${h - 2} L ${(3 * w) / 4},2`;
		case 'textFadeLeft':
			return `M 2,${h / 2} L ${w - 2},2 L ${w - 2},${h - 2} Z`;
		case 'textFadeRight':
			return `M 2,2 L ${w - 2},${h / 2} L 2,${h - 2} Z`;
		case 'textTriangle':
		case 'textTriangleInverted':
			return `M ${w / 2},2 L ${w - 2},${h - 2} L 2,${h - 2} Z`;
		case 'textChevron':
		case 'textChevronInverted':
			return `M 2,2 L ${w / 2},${h - 2} L ${w - 2},2`;
		case 'textRingInside':
		case 'textRingOutside':
			return `M ${w / 2},2 A ${w / 2 - 2},${h / 2 - 2} 0 1,1 ${w / 2},${h - 2} A ${w / 2 - 2},${h / 2 - 2} 0 1,1 ${w / 2},2`;
		case 'textStop':
			return `M ${w * 0.2},2 L ${w * 0.8},2 L ${w - 2},${h * 0.3} L ${w - 2},${h * 0.7} L ${w * 0.8},${h - 2} L ${w * 0.2},${h - 2} L 2,${h * 0.7} L 2,${h * 0.3} Z`;
		case 'textCanUp':
		case 'textCanDown':
			return `M 2,${h - 2} A ${w / 2 - 2},4 0 0,1 ${w - 2},${h - 2}`;
		default:
			return `M 2,${h / 2} L ${w - 2},${h / 2}`;
	}
}
