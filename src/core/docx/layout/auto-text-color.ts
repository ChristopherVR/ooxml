import { automaticTextColor } from '../../color/contrast';
import type { LayoutParagraph } from './input';

/**
 * Word draws text of the automatic colour white on a dark shading (the paragraph's, else its
 * cell's) and black otherwise; an explicit colour, even black, is kept (#436). Runs without a
 * colour take the white here, so the viewer and both PDF exports agree; black stays the default.
 */
export function withAutoTextColor(paragraph: LayoutParagraph, behind: string | undefined) {
	if (!behind || automaticTextColor(behind) !== '#ffffff') return paragraph;
	if (!paragraph.runs.some((run) => !run.hidden && !run.color)) return paragraph;
	return {
		...paragraph,
		runs: paragraph.runs.map((run) =>
			run.hidden || run.color ? run : { ...run, color: '#ffffff' },
		),
	};
}
