import { twips, twipsFromInches } from '../../units/units';
import type { Paragraph } from '../model';

export const mm = (value: number) => twipsFromInches(value / 25.4);
export const stationeryParagraph = (id: string, text = ''): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [{ text: text.slice(0, 4096).replace(/\r\n?/g, '\n') }],
	spacingBeforeTwips: twips(0),
	spacingAfterTwips: twips(0),
});
