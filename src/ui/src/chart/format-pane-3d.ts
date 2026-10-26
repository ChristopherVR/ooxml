/**
 * The 3-D Format section of a chart Format pane's Effects tab: top and
 * bottom bevel, depth, contour, material and lighting. An extension section
 * (PowerPoint's); bevel and material captions keep PowerPoint's shared keys.
 *
 * @module chart/format-pane-3d
 */
import { hexForInput, t } from './format-pane-common';
import type { ChartFormatChoice, ChartFormatField, ChartFormatSection } from './format-pane-types';
import type { ChartFormatBevelView, ChartFormatThreeDView } from './format-pane-view';

/** `ST_BevelPresetType`, in PowerPoint's gallery order. */
export const CHART_BEVEL_TYPES = [
	'circle',
	'relaxedInset',
	'cross',
	'coolSlant',
	'angle',
	'softRound',
	'convex',
	'slope',
	'divot',
	'riblet',
	'hardEdge',
	'artDeco',
] as const;

const MATERIALS = [
	'warmMatte',
	'matte',
	'plastic',
	'metal',
	'dkEdge',
	'softEdge',
	'flat',
	'softmetal',
	'clear',
	'powder',
	'translucentPowder',
] as const;

/** Lighting rigs offered (`ST_LightRigType`). */
const LIGHTING: Array<[string, string]> = [
	['threePt', 'Three Point'],
	['balanced', 'Balanced'],
	['soft', 'Soft'],
	['harsh', 'Harsh'],
	['flood', 'Flood'],
	['contrasting', 'Contrasting'],
	['flat', 'Flat'],
	['twoPt', 'Two Point'],
	['brightRoom', 'Bright Room'],
];

const bevelChoices = (): ChartFormatChoice[] => [
	{ value: 'none', label: t('marker.none', 'None') },
	...CHART_BEVEL_TYPES.map((type) => ({
		value: type,
		label: t(`bevel.${type}`, type),
	})),
];

const size = (id: string, key: string, fallback: string, value: number): ChartFormatField => ({
	id,
	kind: 'number',
	label: t(key, fallback),
	value: Math.round(value * 100) / 100,
	min: 0,
	max: 1584,
	step: 0.5,
	unit: 'pt',
});

function bevelFields(
	which: 'Top' | 'Bottom',
	bevel: ChartFormatBevelView | null | undefined,
): ChartFormatField[] {
	const fields: ChartFormatField[] = [
		{
			id: `threeD.bevel${which}`,
			kind: 'select',
			label: which === 'Top' ? t('bevelTop', 'Top bevel') : t('bevelBottom', 'Bottom bevel'),
			value: bevel?.type ?? 'none',
			choices: bevelChoices(),
		},
	];
	if (bevel) {
		fields.push(
			size(`threeD.bevel${which}Width`, 'width', 'Width', bevel.width),
			size(`threeD.bevel${which}Height`, 'height', 'Height', bevel.height),
		);
	}
	return fields;
}

/** The 3-D Format section for an element with `threeD` as its current 3-D settings. */
export function threeDSection(threeD: ChartFormatThreeDView): ChartFormatSection {
	return {
		id: 'threeD',
		title: t('section.threeD', '3-D Format'),
		fields: [
			...bevelFields('Top', threeD.bevelTop),
			...bevelFields('Bottom', threeD.bevelBottom),
			{
				id: 'threeD.depthColor',
				kind: 'color',
				label: t('depthColor', 'Depth color'),
				value: hexForInput(threeD.depth?.color),
			},
			size('threeD.depthSize', 'depthSize', 'Depth size', threeD.depth?.size ?? 0),
			{
				id: 'threeD.contourColor',
				kind: 'color',
				label: t('contourColor', 'Contour color'),
				value: hexForInput(threeD.contour?.color),
			},
			size('threeD.contourSize', 'contourSize', 'Contour size', threeD.contour?.size ?? 0),
			{
				id: 'threeD.material',
				kind: 'select',
				label: t('material', 'Material'),
				value: threeD.material ?? 'warmMatte',
				choices: MATERIALS.map((material) => ({
					value: material,
					label: t(`material.${material}`, material),
				})),
			},
			{
				id: 'threeD.lighting',
				kind: 'select',
				label: t('lighting', 'Lighting'),
				value: threeD.lighting ?? 'threePt',
				choices: LIGHTING.map(([value, fallback]) => ({
					value,
					label: t(`lighting.${value}`, fallback),
				})),
			},
		],
		note: t(
			'threeDSaveNote',
			'3-D settings are saved for Office; this preview does not reproduce every 3-D effect.',
		),
	};
}
