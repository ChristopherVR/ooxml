import { WORD_ART_STYLES } from 'ooxml-core/drawingml';
import { drawingSchemeColors, type Workbook } from 'ooxml-core/xlsx';

const accents = [1, 2, 3, 4, 5, 6] as const;
const svg = '<svg xmlns="http://www.w3.org/2000/svg"';
const defaultAccents = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
export const shapeFormatStyles = accents.flatMap((accent) => [
	{ color: { accent }, outlined: false, label: 'Accent {n}' },
	{ color: { accent }, outlined: true, label: 'Accent {n} Outline' },
]);
export const pictureFormatStyles = [
	{ label: 'Simple Frame, White', border: 'FFFFFF', shadow: false },
	{ label: 'Simple Frame, Black', border: '000000', shadow: false },
	{ label: 'Drop Shadow', border: null, shadow: true },
];

export function shapeStyleTiles(
	workbook: Workbook | undefined,
	translate: (key: string, vars: { n: number }) => string,
) {
	const theme = drawingSchemeColors(workbook?.theme);
	return shapeFormatStyles.map(({ color: { accent }, outlined, label }, index) => {
		const ink = theme['accent' + accent] ?? defaultAccents[accent - 1];
		return {
			id: String(index),
			label: translate(label, { n: accent }),
			preview: `${svg} width="48" height="28" viewBox="0 0 48 28"><rect x="4" y="4" width="40" height="20" rx="2" fill="${outlined ? 'none' : ink}" stroke="${ink}" stroke-width="2"/></svg>`,
		};
	});
}

export const pictureStyleTiles = pictureFormatStyles.map(({ label, border, shadow }, index) => ({
	id: String(index),
	label,
	preview: `${svg} width="48" height="28" viewBox="0 0 48 28">${shadow ? '<rect x="10" y="6" width="32" height="21" fill="#666"/>' : ''}<rect x="7" y="3" width="32" height="21" fill="#748FAD" stroke="${border ? '#' + border : 'none'}" stroke-width="2"/><path d="M9 21 17 12 22 17 28 10 37 21Z" fill="#E4F1F8"/></svg>`,
}));

export const wordArtStyleTiles = (translate: (key: string, vars: { n: number }) => string) =>
	WORD_ART_STYLES.map(({ fill, outline, shadow }, index) => ({
		id: String(index),
		label: translate('WordArt style {n}', { n: index + 1 }),
		preview: `${svg} width="52" height="30" viewBox="0 0 52 30">${shadow ? `<text x="27" y="24" text-anchor="middle" font-size="20" font-weight="bold" fill="#${shadow}">Ab</text>` : ''}<text x="25" y="22" text-anchor="middle" font-size="20" font-weight="bold" fill="#${fill}" stroke="${outline ? '#' + outline : 'none'}" paint-order="stroke">Ab</text></svg>`,
	}));
