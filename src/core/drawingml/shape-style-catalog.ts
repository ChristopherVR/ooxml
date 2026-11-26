/**
 * PowerPoint's Shape Styles gallery, as PowerPoint itself writes it.
 *
 * Ground truth: `scripts/capture-gallery-writes-com.ps1` sets
 * `Shape.ShapeStyle = 1..77` (msoShapeStylePreset1..77) on a fresh deck and
 * saves it. Presets 1-42 ("Theme Styles", 6 rows x 7 colours) are a bare
 * `<p:style>` with an empty `spPr`; presets 43-77 ("Presets") author the fill
 * and outline in `spPr` and carry an all-zero `<p:style>` whose `fontRef`
 * picks the text colour. The tables below are that capture, verbatim.
 *
 */
import type { CompleteStyleRefs } from './shape-style-refs';
import type { DrawingColor, DrawingFill, DrawingLine } from './types';

/** The seven colour columns: Dark 1, then Accent 1-6. */
export const SHAPE_STYLE_COLUMNS = [
	'dk1',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
] as const;
export type ShapeStyleColumn = (typeof SHAPE_STYLE_COLUMNS)[number];

/** Row captions of the Theme Styles block (PowerPoint's tooltips). */
export const SHAPE_THEME_STYLE_ROWS = [
	{ key: 'coloredOutline', label: 'Colored Outline' },
	{ key: 'coloredFill', label: 'Colored Fill' },
	{ key: 'light1OutlineColoredFill', label: 'Light 1 Outline, Colored Fill' },
	{ key: 'subtleEffect', label: 'Subtle Effect' },
	{ key: 'moderateEffect', label: 'Moderate Effect' },
	{ key: 'intenseEffect', label: 'Intense Effect' },
] as const;

/** Row captions of the Presets block. */
export const SHAPE_PRESET_STYLE_ROWS = [
	{ key: 'transparentText', label: 'Transparent, Colored Text' },
	{ key: 'transparentOutline', label: 'Transparent, Colored Outline' },
	{ key: 'semitransparentFill', label: 'Semitransparent Fill' },
	{ key: 'solidFill', label: 'Solid Fill' },
	{ key: 'gradientFill', label: 'Gradient Fill' },
] as const;

type RefColor = 'col' | 'lt1' | { shade: number };
interface ThemeRow {
	ln: [number, RefColor];
	fill: [number, 'col' | 'lt1'];
	effect: number;
	font: 'dk1' | 'lt1';
}

const THEME_ROWS: readonly ThemeRow[] = [
	{ ln: [2, 'col'], fill: [1, 'lt1'], effect: 0, font: 'dk1' },
	{ ln: [2, { shade: 15000 }], fill: [1, 'col'], effect: 0, font: 'lt1' },
	{ ln: [3, 'lt1'], fill: [1, 'col'], effect: 1, font: 'lt1' },
	{ ln: [1, 'col'], fill: [2, 'col'], effect: 1, font: 'dk1' },
	{ ln: [1, 'col'], fill: [3, 'col'], effect: 2, font: 'lt1' },
	{ ln: [0, 'col'], fill: [3, 'col'], effect: 3, font: 'lt1' },
];

function schemeColor(
	value: ShapeStyleColumn | 'lt1',
	transforms: DrawingColor['transforms'] = [],
): DrawingColor {
	return { kind: 'scheme', value, transforms };
}

/** The four theme references authored by a Theme Styles tile. */
export function themeShapeStyle(row: number, column: ShapeStyleColumn) {
	const spec = THEME_ROWS[row];
	if (!spec || !SHAPE_STYLE_COLUMNS.includes(column)) return undefined;
	const [line, lineColor] = spec.ln;
	return {
		line: {
			idx: line,
			color:
				lineColor === 'lt1'
					? schemeColor('lt1')
					: schemeColor(
							column,
							typeof lineColor === 'object'
								? [{ name: 'shade', value: String(lineColor.shade) }]
								: [],
						),
		},
		fill: { idx: spec.fill[0], color: schemeColor(spec.fill[1] === 'col' ? column : 'lt1') },
		effect: { idx: spec.effect, color: schemeColor(column) },
		font: { idx: 'minor', color: schemeColor(spec.font) },
	} satisfies CompleteStyleRefs;
}

/** What a Presets entry authors in `spPr`, in theme-colour terms. */
export interface ShapePresetStyleSpec {
	fill:
		| { kind: 'none' }
		| { kind: 'solid'; alpha?: number }
		| {
				kind: 'gradient';
				stops: ReadonlyArray<{ position: number; lumMod: number; lumOff?: number }>;
				angle: number;
		  };
	/** Outline width in EMU, or 0 for `<a:ln><a:noFill/></a:ln>`. */
	lineWidthEmu: number;
	/** `fontRef` colour: the column colour or Light 1. */
	font: 'col' | 'lt1';
}

export const SHAPE_PRESET_ROWS: readonly ShapePresetStyleSpec[] = [
	{ fill: { kind: 'none' }, lineWidthEmu: 0, font: 'col' },
	{ fill: { kind: 'none' }, lineWidthEmu: 9525, font: 'col' },
	{ fill: { kind: 'solid', alpha: 50000 }, lineWidthEmu: 0, font: 'lt1' },
	{ fill: { kind: 'solid' }, lineWidthEmu: 0, font: 'lt1' },
	{
		fill: {
			kind: 'gradient',
			stops: [
				{ position: 0, lumMod: 67000 },
				{ position: 48000, lumMod: 97000, lumOff: 3000 },
				{ position: 100000, lumMod: 60000, lumOff: 40000 },
			],
			angle: 16200000,
		},
		lineWidthEmu: 0,
		font: 'lt1',
	},
];

export function presetShapeStyle(row: number, column: ShapeStyleColumn) {
	const spec = SHAPE_PRESET_ROWS[row];
	if (!spec || !SHAPE_STYLE_COLUMNS.includes(column)) return undefined;
	const zero: DrawingColor = { kind: 'scrgb', value: '0,0,0', transforms: [] };
	const ref = { idx: 0, color: zero };
	const style = {
		line: ref,
		fill: ref,
		effect: ref,
		font: { idx: 'minor', color: schemeColor(spec.font === 'col' ? column : spec.font) },
	} satisfies CompleteStyleRefs;
	const line: DrawingLine = spec.lineWidthEmu
		? { widthEmu: spec.lineWidthEmu, fill: { kind: 'solid', color: schemeColor(column) } }
		: { fill: { kind: 'none' } };
	let fill: DrawingFill;
	switch (spec.fill.kind) {
		case 'none':
			fill = { kind: 'none' };
			break;
		case 'solid':
			fill = {
				kind: 'solid',
				color: schemeColor(
					column,
					spec.fill.alpha === undefined ? [] : [{ name: 'alpha', value: String(spec.fill.alpha) }],
				),
			};
			break;
		case 'gradient':
			fill = {
				kind: 'gradient',
				angle: spec.fill.angle / 60000,
				scaled: true,
				stops: spec.fill.stops.map((stop) => ({
					position: stop.position / 1000,
					color: schemeColor(column, [
						{ name: 'lumMod', value: String(stop.lumMod) },
						...(stop.lumOff === undefined ? [] : [{ name: 'lumOff', value: String(stop.lumOff) }]),
					]),
				})),
			};
			break;
	}
	return { style, fill, line };
}

export interface ShapeStylePick {
	block: 'theme' | 'preset';
	row: number;
	column: ShapeStyleColumn;
}

/** Gallery item ids are shared by the shape style galleries. */
export function parseShapeStylePick(id: string): ShapeStylePick | undefined {
	const match = /^(theme|preset)-(\d)-(\d)$/u.exec(id);
	if (!match) return undefined;
	const [, block] = match;
	const row = Number(match[2]);
	const column = SHAPE_STYLE_COLUMNS[Number(match[3])];
	if ((block !== 'theme' && block !== 'preset') || !column) return undefined;
	const rows = block === 'theme' ? SHAPE_THEME_STYLE_ROWS : SHAPE_PRESET_STYLE_ROWS;
	return row < rows.length ? { block, row, column } : undefined;
}
