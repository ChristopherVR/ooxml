/** The report's page measures, list definitions and theme. */
import type { ThemeCatalog } from '../../../../src/core/docx/index';
import { FONTS } from '../../lib/fiction';
import { numIdOf, type ListSpec } from '../kit/parts';
import { PALETTE } from './data';

/** Text width of a portrait page and of the landscape section, in twips (2.5 cm margins). */
export const MARGIN = 1418;
export const TEXT_WIDTH = 11906 - 2 * MARGIN;
export const WIDE_WIDTH = 16838 - 2 * MARGIN;

/** Hanging indent of the heading numbers, twips. */
const NUMBER_INDENT = 794;

const HEADINGS: ListSpec = {
	id: 0,
	levels: [
		{ format: 'decimal', text: '%1', style: 'Heading1' },
		{ format: 'decimal', text: '%1.%2', style: 'Heading2' },
		{ format: 'decimal', text: '%1.%2.%3', style: 'Heading3' },
	].map((level) => ({
		...level,
		paragraph: {
			tabs: [{ pos: NUMBER_INDENT, align: 'left' as const }],
			left: NUMBER_INDENT,
			hanging: NUMBER_INDENT,
		},
	})),
};

const BULLETS: ListSpec = {
	id: 1,
	type: 'hybridMultilevel',
	levels: [
		{
			format: 'bullet',
			text: '•',
			paragraph: { left: 425, hanging: 283 },
			run: { font: FONTS.serif, color: PALETTE.teal },
		},
		{
			format: 'bullet',
			text: '–',
			paragraph: { left: 794, hanging: 283 },
			run: { font: FONTS.serif, color: PALETTE.teal },
		},
	],
};

const STEPS: ListSpec = {
	id: 2,
	type: 'hybridMultilevel',
	levels: [
		{
			format: 'decimal',
			text: '%1.',
			paragraph: { left: 425, hanging: 425 },
			run: { font: FONTS.body, bold: true, color: PALETTE.teal, size: 9.5 },
		},
	],
};

export const LISTS = [HEADINGS, BULLETS, STEPS];
export const NUM = {
	headings: numIdOf(HEADINGS),
	bullets: numIdOf(BULLETS),
	steps: numIdOf(STEPS),
} as const;

export const THEME: ThemeCatalog = {
	name: 'Saltmarsh',
	colorSchemeName: 'Saltmarsh',
	fontSchemeName: 'Saltmarsh',
	colors: {
		dk1: PALETTE.ink,
		lt1: PALETTE.white,
		dk2: PALETTE.deep,
		lt2: PALETTE.mist,
		accent1: PALETTE.teal,
		accent2: PALETTE.amber,
		accent3: PALETTE.sea,
		accent4: PALETTE.rust,
		accent5: PALETTE.slate,
		accent6: '7A9E7E',
		hlink: PALETTE.teal,
		folHlink: PALETTE.slate,
	},
	colorMapping: {},
	fonts: {
		major: { latin: FONTS.display, eastAsia: FONTS.body, complexScript: FONTS.body },
		minor: { latin: FONTS.serif, eastAsia: FONTS.body, complexScript: FONTS.body },
	},
};
