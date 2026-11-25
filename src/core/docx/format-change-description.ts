/**
 * Describes a tracked formatting change (`w:rPrChange`, `w:pPrChange`) the way Word's markup
 * balloons and Reviewing Pane word it: "Formatted: Font: Bold, Font color: Red", "Formatted:
 * Centered, Space After: 6 pt". The prior properties come from the revision's snapshot; the
 * description lists the current values of the direct properties that differ. Measurements are
 * written in points (spacing) and inches (indents), as Word does with US units.
 */
import type { Paragraph, ParagraphStyleCatalog, TextRun } from './model';
import type { RunStyleCatalog } from './run-style-model';
import type { ThemeColorReference } from './theme-model';
import { restoreParagraphFormatting, PARAGRAPH_FORMAT_KEYS } from './restore-paragraph-format';
import { restoreRunFormatting } from './restore-run-format';
import { DIRECT_RUN_PROPERTY_KEYS } from './run-formatting';

export interface FormatChangeStyles {
	paragraphStyles?: ParagraphStyleCatalog;
	characterStyles?: RunStyleCatalog;
}

const NAMED_COLORS: Record<string, string> = {
	FF0000: 'Red',
	C00000: 'Dark Red',
	FFC000: 'Orange',
	FFFF00: 'Yellow',
	'92D050': 'Light Green',
	'00B050': 'Green',
	'00B0F0': 'Light Blue',
	'0070C0': 'Blue',
	'002060': 'Dark Blue',
	'7030A0': 'Purple',
	'000000': 'Black',
	FFFFFF: 'White',
};

const THEME_COLORS: Record<string, string> = {
	text1: 'Text 1',
	dark1: 'Text 1',
	background1: 'Background 1',
	light1: 'Background 1',
	text2: 'Text 2',
	dark2: 'Text 2',
	background2: 'Background 2',
	light2: 'Background 2',
	hyperlink: 'Hyperlink',
	followedHyperlink: 'Followed Hyperlink',
};

const number = (value: number) => String(Number(value.toFixed(2)));
const points = (twips: number) => `${number(twips / 20)} pt`;
const inches = (twips: number) => `${number(twips / 1440)}"`;
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const title = (value: string) =>
	value.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/^./, (c) => c.toUpperCase());

/** A colour as Word names it: a standard colour, a theme colour, or `Custom Color(RGB(...))`. */
export function wordColorName(color: string | undefined, theme?: ThemeColorReference): string {
	if (theme) {
		const base =
			THEME_COLORS[theme.token] ??
			(/^accent(\d)$/.exec(theme.token) ? `Accent ${theme.token.slice(6)}` : title(theme.token));
		if (theme.tint !== undefined && theme.tint < 1)
			return `${base}, Lighter ${Math.round((1 - theme.tint) * 100)}%`;
		if (theme.shade !== undefined && theme.shade < 1)
			return `${base}, Darker ${Math.round((1 - theme.shade) * 100)}%`;
		return base;
	}
	if (!color || color === 'auto') return 'Auto';
	const hex = color.replace(/^#/, '').toUpperCase();
	if (NAMED_COLORS[hex]) return NAMED_COLORS[hex];
	if (!/^[\dA-F]{6}$/.test(hex)) return color;
	const channel = (index: number) => Number.parseInt(hex.slice(index, index + 2), 16);
	return `Custom Color(RGB(${channel(0)},${channel(2)},${channel(4)}))`;
}

const toggle = (value: boolean | undefined, on: string, off = `Not ${on}`) => (value ? on : off);

/** The font items ("Bold", "14 pt") and the other items ("Font color: Red") of a run change. */
function runItems(before: TextRun, after: TextRun, styles: FormatChangeStyles): string[] {
	const changed = (key: keyof TextRun) => !same(before[key], after[key]);
	const font: string[] = [];
	const other: string[] = [];
	if (changed('style')) {
		const name = after.style
			? (styles.characterStyles?.styles[after.style]?.name ?? after.style)
			: 'Default Paragraph Font';
		other.push(name);
	}
	if (changed('fontFamily') || changed('fontTheme'))
		font.push(
			after.fontFamily ??
				(after.fontTheme?.ascii ? `+${title(after.fontTheme.ascii)} (Theme font)` : '(Default)'),
		);
	if (changed('fontSize') && after.fontSize !== undefined)
		font.push(`${number(after.fontSize)} pt`);
	if (changed('bold')) font.push(toggle(after.bold, 'Bold'));
	if (changed('italic')) font.push(toggle(after.italic, 'Italic'));
	if (changed('underline') || changed('underlineStyle')) {
		const style = after.underlineStyle;
		font.push(
			style === 'none' || (!after.underline && !style)
				? 'No underline'
				: !style || style === 'single'
					? 'Underline'
					: `${title(style)} underline`,
		);
	}
	if (changed('underlineColor'))
		font.push(`Underline color: ${wordColorName(after.underlineColor)}`);
	if (changed('strike')) font.push(toggle(after.strike, 'Strikethrough'));
	if (changed('doubleStrike')) font.push(toggle(after.doubleStrike, 'Double strikethrough'));
	if (changed('verticalAlign'))
		font.push(
			after.verticalAlign === 'superscript'
				? 'Superscript'
				: after.verticalAlign === 'subscript'
					? 'Subscript'
					: 'Not Superscript/ Subscript',
		);
	if (changed('caps')) font.push(toggle(after.caps, 'All caps'));
	if (changed('smallCaps')) font.push(toggle(after.smallCaps, 'Small caps'));
	if (changed('vanish')) font.push(toggle(after.vanish, 'Hidden'));
	if (changed('boldComplexScript')) font.push(toggle(after.boldComplexScript, 'Bold (Complex)'));
	if (changed('italicComplexScript'))
		font.push(toggle(after.italicComplexScript, 'Italic (Complex)'));
	if (changed('fontFamilyComplexScript') && after.fontFamilyComplexScript)
		font.push(`(Complex) ${after.fontFamilyComplexScript}`);
	if (changed('fontFamilyEastAsia') && after.fontFamilyEastAsia)
		font.push(`(Asian) ${after.fontFamilyEastAsia}`);
	if (changed('fontSizeComplexScript') && after.fontSizeComplexScript !== undefined)
		font.push(`Complex Script Font: ${number(after.fontSizeComplexScript)} pt`);
	if (changed('color') || changed('colorTheme'))
		other.push(`Font color: ${wordColorName(after.color, after.colorTheme)}`);
	if (changed('highlight')) other.push(after.highlight ? 'Highlight' : 'Not Highlight');
	if (changed('characterSpacingTwips')) {
		const twips = after.characterSpacingTwips ?? 0;
		other.push(
			twips > 0
				? `Expanded by ${points(twips)}`
				: twips < 0
					? `Condensed by ${points(-twips)}`
					: 'Character spacing: Normal',
		);
	}
	if (changed('textScalePercent')) other.push(`Character scale: ${after.textScalePercent ?? 100}%`);
	if (changed('kerningHalfPoints'))
		other.push(
			after.kerningHalfPoints ? `Kern at ${number(after.kerningHalfPoints / 2)} pt` : 'Not Kern',
		);
	if (changed('positionHalfPoints')) {
		const half = after.positionHalfPoints ?? 0;
		other.push(
			half > 0
				? `Raised by ${number(half / 2)} pt`
				: half < 0
					? `Lowered by ${number(-half / 2)} pt`
					: 'Position: Normal',
		);
	}
	if (changed('shadingFill') || changed('shadingThemeFill'))
		other.push(
			after.shadingFill || after.shadingThemeFill
				? `Pattern: Clear (${wordColorName(after.shadingFill, after.shadingThemeFill)})`
				: 'Pattern: Clear',
		);
	if (changed('language') || changed('eastAsiaLanguage') || changed('bidiLanguage'))
		other.push(
			`Language: ${after.language ?? after.eastAsiaLanguage ?? after.bidiLanguage ?? 'Default'}`,
		);
	if (changed('rtl')) other.push(toggle(after.rtl, 'Right-to-left', 'Left-to-right'));
	if (changed('ligatures')) other.push('Ligatures');
	const described = font.length || other.length;
	if (!described && DIRECT_RUN_PROPERTY_KEYS.some((key) => changed(key))) other.push('Font');
	return [...(font.length ? [`Font: ${font.join(', ')}`] : []), ...other];
}

const ALIGNMENTS: Record<string, string> = {
	left: 'Left',
	start: 'Left',
	center: 'Centered',
	right: 'Right',
	end: 'Right',
	justify: 'Justified',
	both: 'Justified',
	distribute: 'Distributed',
};

function lineSpacing(paragraph: Paragraph): string {
	const value = paragraph.lineSpacingTwips;
	if (value === undefined) return 'single';
	if (paragraph.lineSpacingRule === 'exact') return `Exactly ${points(value)}`;
	if (paragraph.lineSpacingRule === 'atLeast') return `At least ${points(value)}`;
	if (value === 240) return 'single';
	if (value === 360) return '1.5 lines';
	if (value === 480) return 'double';
	return `Multiple ${number(value / 240)} li`;
}

function paragraphItems(before: Paragraph, after: Paragraph, styles: FormatChangeStyles): string[] {
	const changed = (key: keyof Paragraph) => !same(before[key], after[key]);
	const items: string[] = [];
	if (changed('style'))
		items.push(
			after.style ? (styles.paragraphStyles?.styles[after.style]?.name ?? after.style) : 'Normal',
		);
	if (changed('align') || changed('justification')) {
		const value = after.justification ?? after.align ?? 'left';
		items.push(ALIGNMENTS[value] ?? title(value));
	}
	const indents: string[] = [];
	if (changed('indentLeftTwips') || changed('indentStartTwips'))
		indents.push(`Left: ${inches(after.indentLeftTwips ?? after.indentStartTwips ?? 0)}`);
	if (changed('indentRightTwips') || changed('indentEndTwips'))
		indents.push(`Right: ${inches(after.indentRightTwips ?? after.indentEndTwips ?? 0)}`);
	if (changed('firstLineTwips') || changed('hangingTwips'))
		indents.push(
			after.hangingTwips
				? `Hanging: ${inches(after.hangingTwips)}`
				: `First line: ${inches(after.firstLineTwips ?? 0)}`,
		);
	if (indents.length) items.push(`Indent: ${indents.join(', ')}`);
	if (changed('spacingBeforeTwips') || changed('spacingBeforeAuto'))
		items.push(
			`Space Before: ${after.spacingBeforeAuto ? 'Auto' : points(after.spacingBeforeTwips ?? 0)}`,
		);
	if (changed('spacingAfterTwips') || changed('spacingAfterAuto'))
		items.push(
			`Space After: ${after.spacingAfterAuto ? 'Auto' : points(after.spacingAfterTwips ?? 0)}`,
		);
	if (changed('lineSpacingTwips') || changed('lineSpacingRule'))
		items.push(`Line spacing: ${lineSpacing(after)}`);
	if (changed('direction'))
		items.push(after.direction === 'rtl' ? 'Right-to-left' : 'Left-to-right');
	if (changed('numbering'))
		items.push(
			after.numbering
				? `Numbered + Level: ${after.numbering.level + 1}`
				: 'No bullets or numbering',
		);
	if (changed('tabStops')) {
		const stops = (after.tabStops ?? []).map(
			(stop) =>
				`${inches(stop.posTwips)}, ${title(stop.align)}${stop.leader && stop.leader !== 'none' ? ` + Leader: ${title(stop.leader)}` : ''}`,
		);
		const removed = (before.tabStops ?? [])
			.filter((stop) => !(after.tabStops ?? []).some((next) => next.posTwips === stop.posTwips))
			.map((stop) => `Not at ${inches(stop.posTwips)}`);
		items.push(`Tab stops: ${[...stops, ...removed].join(' + ') || 'None'}`);
	}
	if (changed('keepNext')) items.push(after.keepNext ? 'Keep with next' : "Don't keep with next");
	if (changed('keepLines'))
		items.push(after.keepLines ? 'Keep lines together' : "Don't keep lines together");
	if (changed('pageBreakBefore'))
		items.push(after.pageBreakBefore ? 'Page break before' : 'No page break before');
	if (changed('widowControl'))
		items.push(after.widowControl === false ? 'Allow widows/orphans' : 'Widow/Orphan control');
	if (changed('contextualSpacing'))
		items.push(
			after.contextualSpacing
				? "Don't add space between paragraphs of the same style"
				: 'Add space between paragraphs of the same style',
		);
	if (changed('suppressLineNumbers'))
		items.push(after.suppressLineNumbers ? 'Suppress line numbers' : 'Line numbers');
	if (changed('outlineLevel'))
		items.push(
			after.outlineLevel
				? `Outline level: Level ${after.outlineLevel}`
				: 'Outline level: Body Text',
		);
	if (changed('borders')) items.push(after.borders ? 'Border' : 'No border');
	if (changed('shadingFill'))
		items.push(
			after.shadingFill ? `Pattern: Clear (${wordColorName(after.shadingFill)})` : 'Pattern: Clear',
		);
	if (changed('dropCap')) items.push(after.dropCap ? 'Drop cap' : 'No drop cap');
	if (!items.length && PARAGRAPH_FORMAT_KEYS.some((key) => changed(key))) items.push('Paragraph');
	return items;
}

/**
 * The items of a run's formatting change, or `undefined` when the run has none or its prior
 * properties cannot be read. An empty list means nothing modeled differs.
 */
export function describeRunFormatChange(
	run: TextRun,
	styles: FormatChangeStyles = {},
): string[] | undefined {
	const revision =
		run.formatRevision?.kind === 'formatChange'
			? run.formatRevision
			: run.revision?.kind === 'formatChange'
				? run.revision
				: undefined;
	if (!revision) return undefined;
	const before = { ...run };
	try {
		restoreRunFormatting(before);
	} catch {
		return undefined;
	}
	return runItems(before, run, styles);
}

/** The items of a paragraph's formatting change (see {@link describeRunFormatChange}). */
export function describeParagraphFormatChange(
	paragraph: Paragraph,
	styles: FormatChangeStyles = {},
): string[] | undefined {
	if (paragraph.formatRevision?.kind !== 'paragraphChange') return undefined;
	const before = { ...paragraph };
	try {
		restoreParagraphFormatting(before);
	} catch {
		return undefined;
	}
	return paragraphItems(before, paragraph, styles);
}

/** The balloon text Word shows: `Formatted: <items>`, or `Formatted` alone. */
export const formatChangeText = (items: readonly string[]): string =>
	items.length ? `Formatted: ${items.join(', ')}` : 'Formatted';

export {
	describeCellFormatChange,
	describeRowFormatChange,
	describeSectionFormatChange,
	describeTableFormatChange,
} from './structure-change-description';
