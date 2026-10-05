import type { VisioParagraph, VisioTextRun } from './model.js';
import { number, sectionRows, type Cells, type Report, type Sheet } from './sheet.js';
import type { Resources } from './style.js';

export interface ParagraphMarker {
	offset: number;
	index: string;
}
const BULLETS: Record<number, string> = { 1: '•', 2: '◆', 3: '■', 4: '☐', 5: '❖', 6: '➤', 7: '✓' };
function paragraphStyle(
	cells: Cells,
	fontFamily: string,
	fontSize: number,
	resources: Resources,
	report: Report,
): Omit<VisioParagraph, 'start' | 'end'> {
	const align = number(cells, 'HorzAlign', 1, report),
		spacing = number(cells, 'SpLine', -1.2, report);
	const bulletType = number(cells, 'Bullet', 0, report),
		customBullet = cells.get('BulletStr')?.value;
	const bulletText = bulletType ? customBullet || BULLETS[bulletType] : undefined;
	if (bulletType && !bulletText)
		report('unsupported-bullet', `Bullet style ${bulletType} is unsupported and was omitted.`);
	const fontId = cells.get('BulletFont')?.value;
	const bulletFont =
		fontId && fontId !== '0'
			? (resources.fonts.get(fontId) ?? (!/^\d+$/.test(fontId) ? fontId : fontFamily))
			: fontFamily;
	const bulletSize = number(cells, 'BulletFontSize', -1, report);
	return {
		horizontalAlign:
			align === 1
				? 'center'
				: align === 2
					? 'right'
					: align === 3
						? 'justify'
						: align === 4
							? 'distributed'
							: 'left',
		indentLeft: number(cells, 'IndLeft', 0, report),
		indentRight: number(cells, 'IndRight', 0, report),
		indentFirst: number(cells, 'IndFirst', 0, report),
		spaceBefore: number(cells, 'SpBefore', 0, report),
		spaceAfter: number(cells, 'SpAfter', 0, report),
		lineSpacing: { kind: spacing < 0 ? 'multiple' : 'exact', value: Math.abs(spacing) },
		direction: number(cells, 'Flags', 0, report) === 1 ? 'rtl' : 'ltr',
		...(bulletText
			? {
					bullet: {
						text: bulletText,
						fontFamily: bulletFont,
						fontSize: bulletSize > 0 ? bulletSize : fontSize,
						offset: number(cells, 'TextPosAfterBullet', 0, report),
					},
				}
			: {}),
	};
}
/** Keep paragraph offsets separate from runs so formatting markers never become displayed text. */
export function textParagraphs(
	text: string,
	markers: ParagraphMarker[],
	sheet: Sheet,
	fontFamily: string,
	fontSize: number,
	resources: Resources,
	report: Report,
	runs: VisioTextRun[] = [],
	consume: () => void = () => {},
): VisioParagraph[] {
	if (!text) return [];
	const rows = new Map(sectionRows(sheet, 'Paragraph').map((row) => [row.index, row.cells]));
	const defaults = rows.get('0') ?? new Map();
	const paragraphs: VisioParagraph[] = [];
	let start = 0,
		markerIndex = 0,
		styleIndex = '0',
		runIndex = 0,
		runStart = 0;
	const append = (end: number) => {
		consume();
		while (markerIndex < markers.length && markers[markerIndex]!.offset <= start)
			styleIndex = markers[markerIndex++]!.index;
		const style = rows.get(styleIndex);
		if (!style && styleIndex !== '0')
			report('missing-paragraph-style', `Paragraph style ${styleIndex} could not be resolved.`);
		while (runIndex < runs.length - 1 && runStart + runs[runIndex]!.text.length <= start)
			runStart += runs[runIndex++]!.text.length;
		const firstRun = runs[runIndex];
		paragraphs.push({
			start,
			end,
			...paragraphStyle(
				style ?? defaults,
				firstRun?.fontFamily ?? fontFamily,
				firstRun?.fontSize ?? fontSize,
				resources,
				report,
			),
		});
	};
	for (const match of text.matchAll(/\r\n|\r|\n/g)) {
		append(match.index);
		start = match.index + match[0].length;
	}
	if (start < text.length) append(text.length);
	return paragraphs;
}
