/** The reference documents, the subset font files they are exported with and Word's lines. */
import { readFileSync } from 'node:fs';
import type { PdfFontFile } from '../../../../pdf/write/pdf-embedded-font';
import words from './word-lines.json';

export const REFERENCES = [
	'letter',
	'table-header',
	'table-header-no-widow-control',
	'header-footer',
	'footnotes',
	'landscape',
] as const;
export type ReferenceName = (typeof REFERENCES)[number];

/** Word 16.0.20625's Save as PDF of each reference, as `pdfLines` reads it. */
export const WORD_LINES: Readonly<Record<ReferenceName, string[][]>> = words;

const file = (name: string) =>
	new Uint8Array(readFileSync(new URL(`../fonts/${name}.subset.ttf`, import.meta.url)));

/** The subsets of the free fonts the documents name (Latin letters, digits and punctuation). */
export function referenceFonts(): PdfFontFile[] {
	return [
		{ family: 'Carlito', data: file('FixtureMetricSans-Regular') },
		{ family: 'Noto Serif', data: file('NotoSerif-Regular') },
		{ family: 'Liberation Sans', data: file('LiberationSans-Regular') },
		{ family: 'Liberation Sans', bold: true, data: file('LiberationSans-Bold') },
		{ family: 'Liberation Serif', data: file('LiberationSerif-Regular') },
	];
}

export const referenceDocx = (name: ReferenceName) =>
	new Uint8Array(readFileSync(new URL(`./${name}.docx`, import.meta.url)));
