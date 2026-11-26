import {
	PICTURE_STYLES,
	SHAPE_STYLE_COLUMNS,
	SHAPE_THEME_STYLE_ROWS,
	SHAPE_PRESET_STYLE_ROWS,
	themeShapeStyle,
	presetShapeStyle,
	WORD_ART_STYLES,
} from 'ooxml-core/drawingml';
import { shapeSvg, type Workbook } from 'ooxml-core/xlsx';
import { pictureStyleTileSvg } from '../../ribbon/picture-style-tile-svg';

const svg = '<svg xmlns="http://www.w3.org/2000/svg"';

export function shapeStyleTiles(
	workbook: Workbook | undefined,
	translate: (key: string, vars?: { n: number }) => string,
) {
	const tiles = (block: 'theme' | 'preset', rows: readonly { label: string }[], section: string) =>
		rows.flatMap(({ label }, row) =>
			SHAPE_STYLE_COLUMNS.map((column, index) => {
				const id = `${block}-${row}-${index}`;
				const style =
					block === 'theme'
						? { style: themeShapeStyle(row, column)! }
						: presetShapeStyle(row, column);
				const preview = shapeSvg(
					{
						kind: 'shape',
						geometry: 'rect',
						frame: { x: 0, y: 0, width: 0, height: 0 },
						...style,
					},
					44,
					26,
					{ theme: workbook?.theme, idPrefix: `xlsx-style-${id}-` },
				);
				const color = index === 0 ? translate('Dark 1') : translate('Accent {n}', { n: index });
				return { id, section, label: `${translate(label)} - ${color}`, preview };
			}),
		);
	return [
		...tiles('theme', SHAPE_THEME_STYLE_ROWS, 'Theme Styles'),
		...tiles('preset', SHAPE_PRESET_STYLE_ROWS, 'Presets'),
	];
}

export const pictureStyleTiles = PICTURE_STYLES.map((spec) => ({
	id: spec.key,
	label: spec.label,
	preview: pictureStyleTileSvg(spec),
}));

export const wordArtStyleTiles = (translate: (key: string, vars: { n: number }) => string) =>
	WORD_ART_STYLES.map(({ fill, outline, shadow }, index) => ({
		id: String(index),
		label: translate('WordArt style {n}', { n: index + 1 }),
		preview: `${svg} width="52" height="30" viewBox="0 0 52 30">${shadow ? `<text x="27" y="24" text-anchor="middle" font-size="20" font-weight="bold" fill="#${shadow}">Ab</text>` : ''}<text x="25" y="22" text-anchor="middle" font-size="20" font-weight="bold" fill="#${fill}" stroke="${outline ? '#' + outline : 'none'}" paint-order="stroke">Ab</text></svg>`,
	}));
