/** The palette and the text styles the three drawings share, so they read as one family. */
import { FONTS } from '../../lib/fiction';
import type { StyleDef, TextStyle } from './styles';

export const INK = '#1E2A32';
export const SLATE = '#51616D';
export const STONE = '#8A98A3';
export const MIST = '#D5DEE3';
export const FOG = '#EEF2F4';
export const PAPER = '#FFFFFF';

export interface Hue {
	deep: string;
	base: string;
	tint: string;
}
export const HUES = {
	teal: { deep: '#0B5E57', base: '#0F766E', tint: '#D7EDEA' },
	sea: { deep: '#0A4259', base: '#0B5A7A', tint: '#D9E9F1' },
	amber: { deep: '#8A5A00', base: '#D98E04', tint: '#FBECC9' },
	coral: { deep: '#8E2F1F', base: '#C8553D', tint: '#F7DDD6' },
	moss: { deep: '#3F6426', base: '#5B8C3A', tint: '#E1EDD6' },
	plum: { deep: '#4B3266', base: '#6D4C8D', tint: '#E7DFF0' },
} as const satisfies Record<string, Hue>;
export type HueName = keyof typeof HUES;

/** Character formats for the runs of mixed text. */
export const RUN = {
	strong: { font: FONTS.display, size: 9, color: INK },
	name: { font: FONTS.display, size: 8.5, color: INK },
	step: { font: FONTS.body, size: 8.5, color: INK },
	body: { font: FONTS.body, size: 8, color: SLATE },
	small: { font: FONTS.body, size: 7, color: SLATE },
	value: { font: FONTS.body, size: 9, color: INK },
} as const satisfies Record<string, TextStyle>;

/** Styles every drawing has: page furniture, labels and the neutral frames. */
export const SHARED_STYLES: readonly StyleDef[] = [
	{
		name: 'Title',
		based: 'Text Only',
		text: { font: FONTS.display, size: 22, color: INK, align: 0, valign: 1, margin: 0 },
	},
	{
		name: 'Subtitle',
		based: 'Text Only',
		text: { font: FONTS.body, size: 10, color: SLATE, align: 0, valign: 1, margin: 0 },
	},
	{
		name: 'Caption',
		based: 'Text Only',
		text: { font: FONTS.body, size: 8, color: SLATE, align: 0, valign: 1, margin: 0 },
	},
	{
		name: 'Body',
		based: 'Caption',
		text: { valign: 0, leading: 1.3 },
	},
	{
		name: 'Caption right',
		based: 'Caption',
		text: { align: 2 },
	},
	{
		name: 'Overline',
		based: 'Text Only',
		text: {
			font: FONTS.display,
			size: 7.5,
			color: STONE,
			align: 0,
			valign: 1,
			margin: 0,
			caps: true,
			tracking: 0.6,
		},
	},
	{
		name: 'Block value',
		based: 'Text Only',
		text: { font: FONTS.body, size: 9, color: INK, align: 0, valign: 1, margin: 0 },
	},
	{
		name: 'Hairline',
		line: { pattern: 0 },
		fill: { color: MIST },
	},
	{
		name: 'Heading',
		based: 'Text Only',
		text: { font: FONTS.display, size: 11, color: INK, align: 0, valign: 1, margin: 0 },
	},
	{
		name: 'Fact',
		based: 'Text Only',
		text: { font: FONTS.body, size: 8, color: INK, align: 0, valign: 1, margin: 0 },
	},
	{
		name: 'Fact right',
		based: 'Fact',
		text: { color: SLATE, align: 2 },
	},
	{
		name: 'Panel',
		line: { color: MIST, weight: 0.75, rounding: 0.08 },
		fill: { color: PAPER },
		text: { font: FONTS.body, size: 8, color: SLATE, align: 0, valign: 0, margin: 6 },
	},
];

/** A filled swatch style per hue, used by legends and accents. */
export function swatchStyles(): StyleDef[] {
	return (Object.keys(HUES) as HueName[]).map((name) => ({
		name: `Swatch ${name}`,
		line: { pattern: 0 },
		fill: { color: HUES[name].base },
	}));
}
