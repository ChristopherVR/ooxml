/**
 * Issue #141's slides 6 -> 7, 9 -> 10 and 17 -> 18, rebuilt by hand as the
 * morph plan sees them, so the morph facts the issue turned on stay pinned
 * without committing the reporter's deck.
 *
 * Everything the matcher and the plan read is kept: every element of each
 * slide in its stacking order (the two master elements first, as the loader
 * puts them), names, `p:cNvPr` ids, which shapes share an `a16:creationId`
 * (numbered here), geometry, rotation and flips, which pictures share a media
 * part, picture transparency, run fonts, sizes, weights and colour alpha, and
 * which paragraphs repeat across boxes and slides. Every word is replaced
 * with filler (one filler per distinct paragraph, so repeats still repeat),
 * media paths are renumbered, and no picture carries image data.
 *
 * @module render/__fixtures__/morph-issue-141
 */
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

/** One run: its text, font family, size (pt), colour, and bold / colour alpha. */
type Run = readonly [
	text: string,
	fontFamily: string | null,
	fontSize: number | null,
	color: string | null,
	flags?: { readonly bold?: boolean; readonly alpha?: number },
];

type Box = readonly [x: number, y: number, width: number, height: number];

/**
 * A text box's own style: its font, size, alignment and colour, whether its
 * body shrinks text on overflow (square wrap), and its insets (null = unset).
 */
function box(
	fontFamily: string,
	fontSize: number,
	align: string,
	color: string,
	shrink: boolean,
	insets: readonly (number | null)[] | null,
): Readonly<Record<string, unknown>> {
	const [left, top, right, bottom] = insets ?? [];
	return {
		fontFamily,
		fontSize,
		align,
		color,
		...(shrink ? { autoFit: true, autoFitMode: 'shrink', textWrap: 'square' } : {}),
		...(left === null || left === undefined ? {} : { bodyInsetLeft: left }),
		...(top === null || top === undefined ? {} : { bodyInsetTop: top }),
		...(right === null || right === undefined ? {} : { bodyInsetRight: right }),
		...(bottom === null || bottom === undefined ? {} : { bodyInsetBottom: bottom }),
	};
}

/**
 * One of the deck's "L" columns: twelve one-letter 72 pt paragraphs in white,
 * each at its own colour alpha (null = opaque), split by 24 pt breaks and
 * closed by an empty 16 pt paragraph.
 */
function letterColumn(fontFamily: string, alphas: readonly (number | null)[]): Run[] {
	return [
		...alphas.flatMap((alpha): Run[] => [
			alpha === null
				? ['L', fontFamily, 72, '#FFFFFF']
				: ['L', fontFamily, 72, '#FFFFFF', { alpha }],
			['\n', 'Calibri', 24, '#3D4146'],
		]),
		['', 'Calibri', 16, '#3D4146'],
	];
}

/** The `a16:creationId` a fixture number stands for; 0 means the shape has none. */
function creationXml(nvKey: string, creation: number): Record<string, unknown> {
	if (creation === 0) {
		return {};
	}
	const guid = `{00000000-0000-4000-8000-${String(creation).padStart(12, '0')}}`;
	return {
		rawXml: {
			[nvKey]: { 'p:cNvPr': { 'a:extLst': { 'a:ext': { 'a16:creationId': { '@_id': guid } } } } },
		},
	};
}

function base(
	id: string,
	name: string,
	shapeId: number,
	[x, y, width, height]: Box,
): Record<string, unknown> {
	return { id, name, shapeId: String(shapeId), x, y, width, height, shapeType: 'rect' };
}

function picture(
	id: string,
	name: string,
	shapeId: number,
	creation: number,
	box: Box,
	media: string,
	extra: Record<string, unknown> = {},
): PptxElement {
	return {
		...base(id, name, shapeId, box),
		type: 'picture',
		imagePath: `ppt/media/${media}`,
		...creationXml('p:nvPicPr', creation),
		...extra,
	} as unknown as PptxElement;
}

function textBody(
	type: 'text' | 'shape',
	id: string,
	name: string,
	shapeId: number,
	creation: number,
	box: Box,
	textStyle: Readonly<Record<string, unknown>>,
	runs: readonly Run[],
	extra: Record<string, unknown> = {},
): PptxElement {
	return {
		...base(id, name, shapeId, box),
		type,
		text: runs.map(([text]) => text).join(''),
		textStyle: { ...textStyle },
		textSegments: runs.map(([text, fontFamily, fontSize, color, flags]) => ({
			text,
			style: {
				...(fontFamily === null ? {} : { fontFamily }),
				...(fontSize === null ? {} : { fontSize }),
				...(color === null ? {} : { color }),
				...(flags?.bold ? { bold: true } : {}),
				...(flags?.alpha === undefined ? {} : { colorRef: { alpha: flags.alpha } }),
			},
		})),
		...creationXml('p:nvSpPr', creation),
		...extra,
	} as unknown as PptxElement;
}

function text(
	id: string,
	name: string,
	shapeId: number,
	creation: number,
	box: Box,
	textStyle: Readonly<Record<string, unknown>>,
	runs: readonly Run[],
	extra?: Record<string, unknown>,
): PptxElement {
	return textBody('text', id, name, shapeId, creation, box, textStyle, runs, extra);
}

function shape(
	id: string,
	name: string,
	shapeId: number,
	creation: number,
	box: Box,
	textStyle: Readonly<Record<string, unknown>>,
	runs: readonly Run[],
	extra?: Record<string, unknown>,
): PptxElement {
	return textBody('shape', id, name, shapeId, creation, box, textStyle, runs, extra);
}

function slide(
	number: string,
	transition: { type: string; durationMs?: number; advanceAfterMs?: number } | null,
	elements: PptxElement[],
): PptxSlide {
	return {
		id: `issue-141-slide-${number}`,
		slideNumber: Number(number),
		elements,
		...(transition ? { transition } : {}),
	} as unknown as PptxSlide;
}

const BOX_1 = box('Calibri', 10.667, 'left', '#3D4146', true, [0, 0, 0, 0]);
const BOX_2 = box('Calibri', 10.667, 'left', '#FFFFFF', true, [0, 0, 0, 0]);
const BOX_3 = box('Calibri Light', 72, 'left', '#FFFFFF', true, [0, 0, 0, 0]);
const BOX_4 = box('Calibri', 72, 'left', '#FFFFFF', true, [0, 0, 0, 0]);
const BOX_5 = box('Calibri', 37.333, 'left', '#FFFFFF', true, null);
const BOX_6 = box('Calibri', 24, 'center', '#FFFFFF', true, null);
const BOX_7 = box('Calibri Light', 24, 'center', '#FFFFFF', false, null);
const BOX_8 = box('Calibri', 37.333, 'left', '#FFFFFF', true, [9.449, null, null, null]);
const BOX_9 = box('Calibri', 18.667, 'left', '#FFFFFF', true, null);

export const SLIDE_6 = slide('6', { type: 'fade', durationMs: 700, advanceAfterMs: 0 }, [
	picture('m6-pic-0', 'Object 11', 11, 1, [1085, 667, 169, 26], 'image1.emf'),
	text(
		'm6-shape-0',
		'Object 7',
		7,
		2,
		[47, 667, 262, 26],
		BOX_1,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#3D4146'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s6-pic-0', 'Object 37', 37, 3, [0, 0, 1280, 720], 'image2.jpeg', {
		flipHorizontal: true,
	}),
	text(
		's6-shape-0',
		'Object 21',
		21,
		4,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's6-shape-1',
		'Object 6',
		6,
		5,
		[1339, 273, 3314, 87],
		BOX_3,
		[
			['ceda', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.09694 }],
			['r j', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.16 }],
			['uniper m', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.29 }],
			['eadow', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.31629 }],
			[' prairie su', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.46 }],
			['mmit', 'Calibri Light', 72, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's6-shape-2',
		'Object 7',
		7,
		6,
		[1666, 361, 4139, 87],
		BOX_3,
		[
			['delta', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.12 }],
			[' kestre', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.14 }],
			['l necta', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.25 }],
			['r quartz t', 'Calibri Light', 72, '#FFFFFF', { alpha: 0.47 }],
			['undra w', 'Calibri Light', 72, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's6-shape-3',
		'Object 14',
		14,
		7,
		[284, -916, 51, 1066],
		BOX_3,
		letterColumn('Calibri Light', [
			0.29747,
			0.29747,
			0.29747,
			0.49,
			0.49,
			null,
			0.49,
			0.27,
			0.07,
			0.24,
			0.24,
			0.24,
		]),
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's6-shape-4',
		'Object 23',
		23,
		8,
		[335, 583, 51, 1066],
		BOX_3,
		letterColumn('Calibri Light', [
			0.29747,
			0.29747,
			0.29747,
			0.49,
			0.49,
			null,
			0.49,
			0.49,
			0.49,
			0.24,
			0.24,
			0.24,
		]),
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's6-shape-5',
		'Object 24',
		24,
		9,
		[385, -941, 144, 1066],
		BOX_3,
		letterColumn('Calibri Light', [
			0.29747,
			0.29747,
			0.29747,
			0.49,
			0.49,
			null,
			0.49,
			0.49,
			0.49,
			0.24,
			0.24,
			0.24,
		]),
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s6-pic-1', 'Object 38', 38, 10, [0, -1, 469, 720], 'image3.jpeg', {
		flipHorizontal: true,
	}),
	picture('s6-pic-2', 'Object 39', 39, 11, [0, 456, 1280, 264], 'image4.jpeg', {
		flipHorizontal: true,
	}),
	picture('s6-pic-3', 'Object 40', 40, 12, [0, -1, 1280, 196], 'image5.jpeg', {
		flipHorizontal: true,
	}),
	text(
		's6-shape-6',
		'Object 46',
		46,
		13,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s6-pic-4', 'Object 3', 3, 14, [1078, 667, 175, 27], 'image6.png', {
		shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 },
	}),
]);
export const SLIDE_7 = slide('7', { type: 'morph', durationMs: 2000, advanceAfterMs: 1000 }, [
	picture('m7-pic-0', 'Object 11', 11, 1, [1085, 667, 169, 26], 'image1.emf'),
	text(
		'm7-shape-0',
		'Object 7',
		7,
		2,
		[47, 667, 262, 26],
		BOX_1,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#3D4146'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s7-pic-0', 'Object 11', 11, 15, [0, 0, 1280, 720], 'image2.jpeg', {
		flipHorizontal: true,
	}),
	text(
		's7-shape-0',
		'Object 21',
		21,
		4,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's7-shape-1',
		'Object 6',
		6,
		5,
		[-2252, 273, 3314, 87],
		BOX_4,
		[
			['ceda', 'Calibri', 72, '#FFFFFF', { alpha: 0.09694 }],
			['r j', 'Calibri', 72, '#FFFFFF', { alpha: 0.16 }],
			['uniper m', 'Calibri', 72, '#FFFFFF', { alpha: 0.29 }],
			['eadow', 'Calibri', 72, '#FFFFFF', { alpha: 0.31629 }],
			[' prairie su', 'Calibri', 72, '#FFFFFF', { alpha: 0.46 }],
			['mmit', 'Calibri', 72, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's7-shape-2',
		'Object 7',
		7,
		6,
		[-2161, 357, 4139, 87],
		BOX_4,
		[
			['delta', 'Calibri', 72, '#FFFFFF', { alpha: 0.12 }],
			[' kestre', 'Calibri', 72, '#FFFFFF', { alpha: 0.14 }],
			['l necta', 'Calibri', 72, '#FFFFFF', { alpha: 0.25 }],
			['r quartz t', 'Calibri', 72, '#FFFFFF', { alpha: 0.47 }],
			['undra w', 'Calibri', 72, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's7-shape-3',
		'Object 3',
		3,
		16,
		[428, 273, 47, 87],
		BOX_4,
		[['L', 'Calibri', 72, '#FFFFFF']],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s7-pic-1', 'Object 12', 12, 17, [0, -3, 469, 720], 'image3.jpeg', {
		flipHorizontal: true,
	}),
	text(
		's7-shape-4',
		'Object 14',
		14,
		7,
		[284, -158, 51, 1066],
		BOX_4,
		letterColumn('Calibri', [
			0.29747,
			0.29747,
			0.29747,
			0.49,
			0.49,
			null,
			0.49,
			0.27,
			0.07,
			0.24,
			0.24,
			0.24,
		]),
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's7-shape-5',
		'Object 23',
		23,
		8,
		[335, -158, 51, 1066],
		BOX_4,
		letterColumn('Calibri', [
			0.29747,
			0.29747,
			0.29747,
			0.49,
			0.49,
			null,
			0.49,
			0.49,
			0.49,
			0.24,
			0.24,
			0.24,
		]),
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's7-shape-6',
		'Object 24',
		24,
		9,
		[385, -157, 144, 1066],
		BOX_4,
		letterColumn('Calibri', [
			0.29747,
			0.29747,
			0.29747,
			0.49,
			0.49,
			null,
			0.49,
			0.49,
			0.49,
			0.24,
			0.24,
			0.24,
		]),
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s7-pic-2', 'Object 15', 15, 18, [0, 456, 1280, 264], 'image7.jpeg', {
		flipHorizontal: true,
		cropRight: -0.00028,
	}),
	picture('s7-pic-3', 'Object 17', 17, 19, [0, -2, 1280, 196], 'image5.jpeg', {
		flipHorizontal: true,
	}),
	text(
		's7-shape-7',
		'Object 18',
		18,
		20,
		[426, 274, 47, 87],
		BOX_4,
		[['L', 'Calibri', 72, '#FFFFFF']],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text(
		's7-shape-8',
		'Object 28',
		28,
		21,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s7-pic-4', 'Object 26', 26, 22, [227, 765, 111, 47], 'image8.png'),
	picture('s7-pic-5', 'Object 29', 29, 23, [685, 1032, 342, 62], 'image9.png'),
	picture('s7-pic-6', 'Object 30', 30, 24, [596, 944, 54, 47], 'image10.png'),
	picture('s7-pic-7', 'Object 31', 31, 25, [380, 854, 177, 46], 'image11.png'),
	picture('s7-pic-8', 'Object 2', 2, 26, [1078, 667, 175, 27], 'image6.png', {
		shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 },
	}),
]);
export const SLIDE_9 = slide('9', { type: 'morph', durationMs: 2000 }, [
	picture('m9-pic-0', 'Object 11', 11, 1, [1085, 667, 169, 26], 'image1.emf'),
	text(
		'm9-shape-0',
		'Object 7',
		7,
		2,
		[47, 667, 262, 26],
		BOX_1,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#3D4146'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s9-pic-0', 'Object 4', 4, 27, [-7, 758, 324, 725], 'image12.jpeg', {
		imageEffects: { alphaModFix: 100 },
	}),
	picture('s9-pic-1', 'Object 18', 18, 28, [0, 0, 1280, 720], 'image2.jpeg', {
		flipHorizontal: true,
	}),
	picture('s9-pic-2', 'Object 2', 2, 29, [0, 0, 1280, 720], 'image13.jpeg', {
		imageEffects: { alphaModFix: 100 },
	}),
	text(
		's9-shape-0',
		'Object 11',
		11,
		30,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s9-pic-3', 'Object 12', 12, 31, [1078, 667, 175, 27], 'image6.png', {
		shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 },
	}),
	picture('s9-pic-4', 'Object 15', 15, 32, [92, 147, 268, 49], 'image14.png'),
	picture('s9-pic-5', 'Object 17', 17, 33, [92, 93, 213, 37], 'image15.png'),
	picture('s9-pic-6', 'Object 20', 20, 34, [-7, 1125, 733, 725], 'image16.png'),
	picture('s9-pic-7', 'Object 21', 21, 35, [611, 1422, 376, 725], 'image17.png'),
	picture('s9-pic-8', 'Object 22', 22, 36, [930, 1837, 350, 725], 'image18.png'),
	text('s9-shape-1', 'Object 24', 24, 37, [47, -304, 660, 55], BOX_5, [
		['ember lagoon orchid ravine', 'Calibri', 37.333, '#FFFFFF'],
	]),
	text('s9-shape-2', 'Object 25', 25, 38, [32, 2383, 229, 39], BOX_6, [
		['fjord meadow', 'Calibri', 24, '#FFFFFF'],
	]),
	text('s9-shape-3', 'Object 26', 26, 39, [329, 2004, 277, 39], BOX_6, [
		['grove nectar quartz tun', 'Calibri', 24, '#FFFFFF'],
	]),
	text('s9-shape-4', 'Object 27', 27, 40, [675, 1736, 229, 39], BOX_6, [
		['heron', 'Calibri', 24, '#FFFFFF'],
	]),
	text('s9-shape-5', 'Object 28', 28, 41, [997, 1276, 229, 39], BOX_6, [
		['inlet p', 'Calibri', 24, '#FFFFFF'],
	]),
]);
export const SLIDE_10 = slide('10', { type: 'morph', durationMs: 2000 }, [
	picture('m10-pic-0', 'Object 11', 11, 1, [1085, 667, 169, 26], 'image1.emf'),
	text(
		'm10-shape-0',
		'Object 7',
		7,
		2,
		[47, 667, 262, 26],
		BOX_1,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#3D4146'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s10-pic-0', 'Object 4', 4, 42, [-7, 2, 324, 723], 'image19.jpeg', {
		imageEffects: { alphaModFix: 100 },
	}),
	picture('s10-pic-1', 'Object 24', 24, 43, [0, 0, 1280, 720], 'image13.jpeg', {
		imageEffects: { alphaModFix: 100 },
	}),
	picture('s10-pic-2', 'Object 5', 5, 44, [-7, 0, 733, 725], 'image16.png'),
	picture('s10-pic-3', 'Object 7', 7, 45, [611, 0, 376, 725], 'image17.png'),
	picture('s10-pic-4', 'Object 9', 9, 46, [930, 0, 350, 725], 'image18.png'),
	picture('s10-pic-5', 'Object 6', 6, 47, [-7, 0, 324, 725], 'image12.jpeg', {
		imageEffects: { alphaModFix: 100 },
	}),
	shape(
		's10-shape-0',
		'Object 23',
		23,
		48,
		[-7, 482, 1287, 243],
		BOX_7,
		[['', 'Calibri Light', 24, '#FFFFFF']],
		{
			rotation: 180,
			shapeStyle: {
				fillMode: 'gradient',
				fillColor: '#1F2023',
				fillOpacity: 0,
				fillGradient: 'linear-gradient(180.00deg, #1F2023 0%, rgba(31, 32, 35, 0) 96%)',
				fillGradientStops: [
					{
						color: '#1F2023',
						position: 0,
						originalColorXml: { 'a:srgbClr': { '@_val': '1F2023' } },
					},
					{
						color: '#1F2023',
						position: 96,
						opacity: 0,
						originalColorXml: { 'a:srgbClr': { 'a:alpha': { '@_val': '0' }, '@_val': '1F2023' } },
					},
				],
				fillGradientAngle: 90,
				fillGradientType: 'linear',
				fillGradientScaled: false,
				lnRefIdx: 2,
				strokeColor: 'transparent',
				strokeWidth: 0,
				strokeDash: 'solid',
				lineJoin: 'miter',
				lineCap: 'flat',
				compoundLine: 'sng',
				inheritedLineStyle: {
					strokeColor: '#002D34',
					strokeWidth: 1.3333333333333333,
					strokeDash: 'solid',
					lineJoin: 'miter',
					lineCap: 'flat',
					compoundLine: 'sng',
				},
				strokeFillMode: 'none',
				inheritedEffectStyle: {},
				fontRefIdx: 'minor',
			},
		},
	),
	shape(
		's10-shape-1',
		'Object 22',
		22,
		49,
		[-7, 0, 1287, 243],
		BOX_7,
		[['', 'Calibri Light', 24, '#FFFFFF']],
		{
			shapeStyle: {
				fillMode: 'gradient',
				fillColor: '#1F2023',
				fillOpacity: 0,
				fillGradient: 'linear-gradient(180.00deg, #1F2023 0%, rgba(31, 32, 35, 0) 96%)',
				fillGradientStops: [
					{
						color: '#1F2023',
						position: 0,
						originalColorXml: { 'a:srgbClr': { '@_val': '1F2023' } },
					},
					{
						color: '#1F2023',
						position: 96,
						opacity: 0,
						originalColorXml: { 'a:srgbClr': { 'a:alpha': { '@_val': '0' }, '@_val': '1F2023' } },
					},
				],
				fillGradientAngle: 90,
				fillGradientType: 'linear',
				fillGradientScaled: false,
				lnRefIdx: 2,
				strokeColor: 'transparent',
				strokeWidth: 0,
				strokeDash: 'solid',
				lineJoin: 'miter',
				lineCap: 'flat',
				compoundLine: 'sng',
				inheritedLineStyle: {
					strokeColor: '#002D34',
					strokeWidth: 1.3333333333333333,
					strokeDash: 'solid',
					lineJoin: 'miter',
					lineCap: 'flat',
					compoundLine: 'sng',
				},
				strokeFillMode: 'none',
				inheritedEffectStyle: {},
				fontRefIdx: 'minor',
			},
		},
	),
	text(
		's10-shape-2',
		'Object 14',
		14,
		0,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#FFFFFF'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	text('s10-shape-3', 'Object 27', 27, 0, [47, 55, 660, 55], BOX_5, [
		['ember lagoon orchid ravine', 'Calibri', 37.333, '#FFFFFF'],
	]),
	text('s10-shape-4', 'Object 21', 21, 50, [32, 559, 229, 39], BOX_6, [
		['fjord meadow', 'Calibri', 24, '#FFFFFF'],
	]),
	text('s10-shape-5', 'Object 28', 28, 51, [329, 559, 277, 39], BOX_6, [
		['grove nectar quartz tun', 'Calibri', 24, '#FFFFFF'],
	]),
	text('s10-shape-6', 'Object 29', 29, 52, [675, 559, 229, 39], BOX_6, [
		['heron', 'Calibri', 24, '#FFFFFF'],
	]),
	text('s10-shape-7', 'Object 30', 30, 53, [997, 561, 229, 39], BOX_6, [
		['inlet p', 'Calibri', 24, '#FFFFFF'],
	]),
	picture('s10-pic-6', 'Object 3', 3, 54, [1078, 667, 175, 27], 'image6.png', {
		shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 },
	}),
]);
export const SLIDE_17 = slide('17', { type: 'cover' }, [
	picture('m17-pic-0', 'Object 11', 11, 1, [1085, 667, 169, 26], 'image1.emf'),
	text(
		'm17-shape-0',
		'Object 7',
		7,
		2,
		[47, 667, 262, 26],
		BOX_1,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#3D4146'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s17-pic-0', 'Object 2', 2, 55, [0, 0, 1280, 720], 'image20.jpeg'),
	text('s17-shape-0', 'Object 20', 20, 0, [75, 91, 330, 459], BOX_8, [
		['juniper qu', 'Calibri', 37.333, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['kestrel ravine upland yarrow basin e', 'Calibri', 24, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['lagoon ', 'Calibri', 16, '#FFFFFF', { bold: true }],
		['summit valley zephyr cedar fj', 'Calibri', 16, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['meado', 'Calibri', 16, '#FFFFFF', { bold: true }],
		['w tundra willow amber delta gro', 'Calibri', 16, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['nectar upland yarrow basin ember her', 'Calibri', 16, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['', 'Calibri', 18.667, '#3D4146'],
	]),
	picture('s17-pic-1', 'Object 4', 4, 56, [640, 0, 640, 720], 'image21.png', {
		imageEffects: { alphaModFix: 0 },
		shapeStyle: {
			shadowColor: '#000000',
			shadowOpacity: 0.4,
			shadowBlur: 57.469711286089236,
			shadowOffsetX: 4,
			shadowOffsetY: 0,
			shadowAngle: 0,
			shadowDistance: 4,
			shadowRotateWithShape: false,
			shadowAlignment: 'l',
			outerShadowOriginalColor: '#000000',
			outerShadowOriginalOpacity: 0.4,
		},
	}),
	picture('s17-pic-2', 'Object 8', 8, 57, [-928, 125, 155, 21], 'image22.png'),
	picture('s17-pic-3', 'Object 9', 9, 58, [-323, 499, 101, 23], 'image23.png'),
	picture('s17-pic-4', 'Object 10', 10, 59, [-830, 172, 264, 31], 'image24.png'),
	picture('s17-pic-5', 'Object 11', 11, 60, [-402, 448, 135, 23], 'image25.png'),
	picture('s17-pic-6', 'Object 12', 12, 61, [-674, 224, 103, 23], 'image26.png'),
	picture('s17-pic-7', 'Object 13', 13, 62, [-623, 269, 109, 23], 'image27.png'),
	picture('s17-pic-8', 'Object 14', 14, 63, [-561, 315, 85, 21], 'image28.png'),
	picture('s17-pic-9', 'Object 15', 15, 64, [-500, 360, 212, 69], 'image29.png'),
	text(
		's17-shape-1',
		'Object 16',
		16,
		65,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['orchid valley zephyr cedar fjord in', 'Calibri', 10.667, '#FFFFFF'],
			['l', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s17-pic-10', 'Object 17', 17, 66, [1078, 667, 175, 27], 'image6.png', {
		shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 },
	}),
	text('s17-shape-2', 'Object 7', 7, 67, [890, 155, 316, 320], BOX_9, [
		['prairie willow amber delta grove jun', 'Calibri', 18.667, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['quartz yarrow basin ember heron kest', 'Calibri', 18.667, '#FFFFFF'],
	]),
	text('s17-shape-3', 'Object 3', 3, 68, [-598, 180, 278, 355], BOX_5, [
		['juniper qu', 'Calibri', 37.333, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['ravine zephyr cedar fjo', 'Calibri', 18.667, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['summit amber delta grove juniper mea', 'Calibri', 21.333, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['', 'Calibri', 18.667, '#3D4146'],
	]),
]);
export const SLIDE_18 = slide('18', { type: 'morph', durationMs: 2000 }, [
	picture('m18-pic-0', 'Object 11', 11, 1, [1085, 667, 169, 26], 'image1.emf'),
	text(
		'm18-shape-0',
		'Object 7',
		7,
		2,
		[47, 667, 262, 26],
		BOX_1,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#3D4146'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['basin inlet lagoon orchid ravine upl', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s18-pic-0', 'Object 2', 2, 55, [0, 0, 1280, 720], 'image20.jpeg'),
	text('s18-shape-0', 'Object 20', 20, 0, [-471, 91, 330, 449], BOX_8, [
		['juniper qu', 'Calibri', 37.333, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['kestrel ravine upland yarrow basin e', 'Calibri', 24, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['tundra basin embe', 'Calibri', 18.667, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['upland cedar fjord inlet lagoon orch', 'Calibri', 18.667, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['valley delta grove juniper m', 'Calibri', 18.667, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['willow ember heron kestrel nectar qu', 'Calibri', 18.667, '#FFFFFF'],
	]),
	text('s18-shape-1', 'Object 7', 7, 67, [1327, 155, 316, 359], BOX_9, [
		['prairie willow amber delta grove jun', 'Calibri', 18.667, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['yarrow fjord inlet lagoon orchid rav', 'Calibri', 18.667, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['nectar upland yarrow basin ember her', 'Calibri', 18.667, '#FFFFFF'],
	]),
	picture('s18-pic-1', 'Object 5', 5, 69, [640, 0, 640, 720], 'image21.png', {
		imageEffects: { alphaModFix: 100 },
		shapeStyle: {
			shadowColor: '#000000',
			shadowOpacity: 0.4,
			shadowBlur: 57.469711286089236,
			shadowOffsetX: 4,
			shadowOffsetY: 0,
			shadowAngle: 0,
			shadowDistance: 4,
			shadowRotateWithShape: false,
			shadowAlignment: 'l',
			outerShadowOriginalColor: '#000000',
			outerShadowOriginalOpacity: 0.4,
		},
	}),
	text('s18-shape-2', 'Object 6', 6, 70, [875, 180, 278, 362], BOX_5, [
		['juniper qu', 'Calibri', 37.333, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['ravine zephyr cedar fjo', 'Calibri', 24, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['summit amber delta grove juniper mea', 'Calibri', 21.333, '#FFFFFF'],
		['\n', 'Calibri', 24, '#3D4146'],
		['', 'Calibri', 18.667, '#3D4146'],
	]),
	picture('s18-pic-2', 'Object 12', 12, 71, [152, 125, 155, 21], 'image22.png'),
	picture('s18-pic-3', 'Object 14', 14, 72, [152, 499, 101, 23], 'image23.png'),
	picture('s18-pic-4', 'Object 16', 16, 73, [152, 172, 264, 31], 'image24.png'),
	picture('s18-pic-5', 'Object 19', 19, 74, [152, 448, 135, 23], 'image25.png'),
	picture('s18-pic-6', 'Object 22', 22, 75, [152, 224, 103, 23], 'image26.png'),
	picture('s18-pic-7', 'Object 24', 24, 76, [152, 269, 109, 23], 'image27.png'),
	picture('s18-pic-8', 'Object 26', 26, 77, [152, 315, 85, 21], 'image28.png'),
	picture('s18-pic-9', 'Object 28', 28, 78, [152, 360, 212, 69], 'image29.png'),
	text(
		's18-shape-3',
		'Object 34',
		34,
		79,
		[47, 667, 262, 26],
		BOX_2,
		[
			['amber heron kestrel nectar quartz tu', 'Calibri', 10.667, '#FFFFFF'],
			['\n', 'Calibri', 10.667, '#E9E3E1'],
			['orchid valley zephyr cedar fjord in', 'Calibri', 10.667, '#FFFFFF'],
			['l', 'Calibri', 10.667, '#3D4146'],
		],
		{ shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 } },
	),
	picture('s18-pic-10', 'Object 35', 35, 80, [1078, 667, 175, 27], 'image6.png', {
		shapeStyle: { fillMode: 'none', fillColor: 'transparent', fillOpacity: 0 },
	}),
]);
