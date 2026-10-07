import type { DiagramTextRun, DiagramTextParagraph } from '../../diagram/types';
import type { ChartStyleEntry } from '../../chart/style-definition';
import type { ChartObject, ThemePalette } from '../model';
import { chartAppearance, type ChartAppearanceEntry } from './chart-appearance';

export interface ChartTitleText {
	paragraphs: (Pick<
		DiagramTextParagraph,
		'align' | 'lineSpacing' | 'spaceBefore' | 'spaceAfter'
	> & { runs: { text: string; appearance: ChartAppearanceEntry }[] })[];
}

function properties(run: Partial<DiagramTextRun>): Partial<ChartStyleEntry> {
	return {
		...(run.sizePt === undefined ? {} : { fontSize: run.sizePt }),
		...(run.bold === undefined ? {} : { bold: run.bold }),
		...(run.italic === undefined ? {} : { italic: run.italic }),
		...(run.underline === undefined ? {} : { underline: run.underline }),
		...(run.typeface === undefined ? {} : { typeface: run.typeface }),
		...(run.color === undefined ? {} : { textColor: run.color }),
	};
}

/** Resolve imported mixed text through the same chart/theme precedence as ordinary labels. */
export function chartTitleText(
	chart: ChartObject,
	theme: ThemePalette,
): ChartTitleText | undefined {
	const formatting = chart.formatting;
	const entry = formatting?.entries.title;
	const body = entry?.textBody;
	if (!formatting || !body || body.text !== chart.title) return undefined;
	if (
		body.paragraphs.length === 1 &&
		body.paragraphs[0]!.runs.length <= 1 &&
		!body.text.includes('\n') &&
		!body.paragraphs[0]!.lineSpacing &&
		!body.paragraphs[0]!.spaceBefore &&
		!body.paragraphs[0]!.spaceAfter
	)
		return undefined;
	return {
		paragraphs: body.paragraphs.map((paragraph) => ({
			...(paragraph.align ? { align: paragraph.align } : {}),
			...(paragraph.lineSpacing ? { lineSpacing: paragraph.lineSpacing } : {}),
			...(paragraph.spaceBefore ? { spaceBefore: paragraph.spaceBefore } : {}),
			...(paragraph.spaceAfter ? { spaceAfter: paragraph.spaceAfter } : {}),
			runs: paragraph.runs.map((run) => {
				const direct = {
					sourceXml: '',
					...entry.textBodyStyle,
					...properties(paragraph.defaultProperties ?? {}),
					...properties(run),
				};
				const appearance =
					chartAppearance(
						{
							...chart,
							formatting: { ...formatting, entries: { ...formatting.entries, title: direct } },
						},
						theme,
					)?.title ?? {};
				return { text: run.text, appearance };
			}),
		})),
	};
}
