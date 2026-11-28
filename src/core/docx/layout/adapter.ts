import { themeLatinFont } from '../default-theme';
import { scriptLayoutRuns } from './script-fonts';
import type {
	Block,
	DocumentModel,
	Paragraph,
	RunFormatting,
	SectionProperties,
	TextRun,
	Twips,
} from '../index';
import {
	computeListLabels,
	dateFieldResult,
	displayListLabel,
	type ParagraphListLabel,
	fieldName,
	isFieldLocked,
	resolveParagraphFormatting,
	resolveRunFormatting,
	resolveThemeColorReference,
	twips,
	twipsToPixels,
} from '../index';
import { floatsOf } from './adapt-floats';
import { equationObject } from './equation-object';
import { shapePaint } from './shape-paint';
import type { TextBoxLayouts } from './text-box-layout';
import { DEFAULT_FONT_SIZE_PT, NO_TWIPS, ptToPx } from './units';
import { docxChartPaint } from '../chart-paint';
import { inkPaint } from './ink-paint';
import { drawingOuterShadow } from '../shape-svg';
import { pictureBulletSizePt } from '../numbering-picture';
import { definedProps } from '../defined-props';
import { contextualStyleId } from '../contextual-spacing';
import { effectiveTextScalePercent } from '../simple-types';
import { kinsokuLanguage } from '../../text/kinsoku';
import { resolveEastAsiaLanguage } from '../run-language';
import { symbolOnly, symbolSegments } from '../special-characters';
import { allCapsText, smallCapsText } from '../caps-display';
import { formFieldDisplayText } from '../form-fields';
import { wordHighlightColors } from '../highlight';
import { paragraphMarkRun } from '../paragraph-mark';
import { isEmptyParagraph } from '../paragraph-content';
import { fontOf } from './paragraph-tokens';
import { adaptTable } from './adapt-table';
import {
	adaptBlockContentControlTags,
	adaptContentControlTags,
} from './adapt-content-control-tags';
import { blockContentControlTags } from './block-content-control-tags';
import type { TableStyleConditionalFormatting } from '../table-model';
import {
	addsParagraphSpacing,
	attachesFootnotesToBodyColumns,
	compatibilityProfile,
	honoursLayoutInCell,
	keepsSpaceBeforePageBreak,
	keepsWrappingFloatsAboveBottomMargin,
	usesModernFootnoteLayout,
	squeezesJustifiedSpaces,
} from '../compatibility';
import { foldDropCaps } from './adapt-drop-cap';
import { adaptTextFrame } from './flow-text-frame';
import { groupParagraphBorders, paragraphBox } from './adapt-paragraph-box';
import { endnoteParagraphs, noteLabels, paragraphFootnotes } from './adapt-notes';
import { isRunHiddenForReview } from '../review-formatting-display';
import {
	visibleReviewModel,
	visibleReviewParagraph,
	reviewSourceLength,
	type LayoutReviewOptions,
} from './review-display';
import type {
	LayoutBlock,
	LayoutDocumentInput,
	LayoutParagraph,
	LayoutParagraphBorders,
	LayoutRun,
	LayoutSection,
} from './input';

/** Word's grey behind form fields while "Form field shading" is on. */
const FORM_FIELD_SHADING = '#d9d9d9';
const TABLE_ROW_NOTE =
	'Table rows with an exact height clip content that does not fit; text in them is not shrunk.';
const VERTICAL_MERGE_NOTE =
	'Vertically merged table cells are drawn as one cell, but their text stays in the first row of the merge.';
const DROP_CAP_NOTE =
	'Drop caps are laid out as an enlarged first letter (a raised cap), not as a dropped cap.';
const TEXT_BOX_NOTE =
	'Text boxes whose shape cannot be drawn are shown on the editing surface but not in Print Layout.';
const EQUATION_NOTE =
	'Equations that cannot be drawn show a marker in Print Layout instead of the typeset equation.';
const NEXT_COLUMN_NOTE = 'A "next column" section break is laid out as a continuous section break.';

// Word 16.0.20625 opens a package without styles using its current Normal template.
const implicitParagraphStyles = {
	docDefaults: {
		lineSpacingTwips: twips(278),
		lineSpacingRule: 'auto',
		spacingAfterTwips: twips(160),
	},
	styles: {},
	warnings: [],
} satisfies NonNullable<DocumentModel['paragraphStyles']>;
const implicitRunStyles = {
	docDefaults: { fontSize: 12, fontFamily: 'Aptos' },
	styles: {},
	warnings: [],
} satisfies NonNullable<DocumentModel['characterStyles']>;

const twipsToPx = twipsToPixels;
/** Word gives `beforeAutospacing` and `afterAutospacing` 14 pt (measured in spacing-probe.json). */
const HTML_AUTO_SPACING_TWIPS = twips(280);

/** A paragraph holding only a display equation (`m:oMathPara`) is centred, as in Word. */
function isDisplayEquationParagraph(paragraph: Paragraph): boolean {
	return (
		paragraph.runs.some((run) => run.equation?.display) &&
		paragraph.runs.every((run) => run.equation?.display || (!run.text && !run.image))
	);
}

function sectionBreak(
	type: SectionProperties['type'],
	note: (message: string) => void,
): NonNullable<LayoutSection['break']> | undefined {
	if (type === 'nextColumn') {
		note(NEXT_COLUMN_NOTE);
		return 'continuous';
	}
	return type;
}

/**
 * Converts a `docx-core` `DocumentModel` into this engine's own input contract. With `textBoxes`,
 * the body's text boxes are laid out by the engine too (`text-box-layout.ts`).
 */
export function adaptDocumentModel(
	model: DocumentModel,
	note: (message: string) => void = () => {},
	options: LayoutReviewOptions = {},
	textBoxes?: TextBoxLayouts,
): LayoutDocumentInput {
	const reviewMode = options.reviewDisplayMode ?? 'all';
	const visibleModel = visibleReviewModel(model, reviewMode);
	const blockTags = blockContentControlTags([
		model.blocks,
		...(visibleModel.footnotes ?? []).map((note) => note.blocks),
		...(visibleModel.endnotes ?? []).map((note) => note.blocks),
		...Object.values(model.footnoteSeparators ?? {}).map((note) => note.blocks),
	]);
	const catalog =
		model.paragraphStyles ?? (model.usesImplicitNormalStyle ? implicitParagraphStyles : undefined);
	const runCatalog =
		model.characterStyles ?? (model.usesImplicitNormalStyle ? implicitRunStyles : undefined);
	const compatibility = compatibilityProfile(model);
	const addParagraphSpacing = addsParagraphSpacing(compatibility);
	const squeeze = squeezesJustifiedSpaces(compatibility);
	const keepSpaceBeforePageBreak = keepsSpaceBeforePageBreak(compatibility);
	const keepWrappingFloatsAboveBottomMargin = keepsWrappingFloatsAboveBottomMargin(compatibility);
	/** HTML auto spacing is 14 pt when the compatibility option is off; otherwise the twips are used. */
	const autoSpaced = (auto: boolean | undefined, twipsValue: Twips | undefined) =>
		auto && !addParagraphSpacing ? HTML_AUTO_SPACING_TWIPS : twipsValue;
	const noted = new Set<string>();
	const reportOnce = (message: string) => {
		if (!noted.has(message)) {
			noted.add(message);
			note(message);
		}
	};

	const now = new Date();
	const theme = model.theme;
	/** Effective run formatting: document defaults, paragraph and character styles, theme fonts. */
	function effective(
		run: TextRun,
		paragraphStyleId: string | undefined,
		tableStyleRun?: RunFormatting,
	) {
		const formatting = resolveRunFormatting(run, {
			runCatalog,
			paragraphCatalog: catalog,
			paragraphStyleId,
			...(tableStyleRun && { tableStyleRun }),
			overrideTableStyleFontSizeAndJustification: model.overrideTableStyleFontSizeAndJustification,
		});
		const role = formatting.fontTheme?.ascii ?? formatting.fontTheme?.hAnsi;
		// Word draws a symbol (`w:sym`) in its own font, whatever the run's font is.
		const family =
			(symbolOnly(run) && run.symbol?.font) ||
			((role ? themeLatinFont(theme, role) : undefined) ?? formatting.fontFamily);
		const color =
			(formatting.colorTheme && theme
				? resolveThemeColorReference(formatting.colorTheme, theme)
				: undefined) ?? formatting.color;
		return { formatting, family, color };
	}
	const noteLabel = noteLabels(visibleModel);
	function adaptRun(
		run: TextRun,
		paragraphStyleId: string | undefined,
		markLabel?: string,
		tableStyleRun?: RunFormatting,
	): LayoutRun {
		if (isRunHiddenForReview(run, reviewMode))
			return { text: run.text, hidden: true, sourceLength: reviewSourceLength(run) };
		const { formatting, family, color } = effective(run, paragraphStyleId, tableStyleRun);
		const scale = effectiveTextScalePercent(formatting.textScalePercent);
		const equation = run.equation
			? equationObject(
					run.equation.omml,
					ptToPx(formatting.fontSize ?? DEFAULT_FONT_SIZE_PT),
					Boolean(run.equation.display),
				)
			: undefined;
		if (run.equation && !equation) reportOnce(EQUATION_NOTE);
		// DATE and TIME update when Word paginates for display or printing.
		const name = run.field ? fieldName(run.field.instr) : '';
		// Note references and a note's own number mark show the note's number, raised.
		const noteText = run.noteReference
			? noteLabel(run.noteReference.kind, run.noteReference.id)
			: run.noteMark
				? (markLabel ?? '')
				: undefined;
		// A legacy check box or drop-down draws its value on its begin marker; Word shades form fields.
		const formText = formFieldDisplayText(run);
		const shaded =
			model.formFieldShading !== false && (formText !== undefined || name === 'FORMTEXT');
		const highlight = formatting.highlight;
		const background =
			highlight && highlight !== 'none'
				? wordHighlightColors[highlight]
				: shaded
					? FORM_FIELD_SHADING
					: undefined;
		const text =
			noteText ??
			formText ??
			(run.field && !isFieldLocked(run) && (name === 'DATE' || name === 'TIME')
				? dateFieldResult(name, run.field.instr, now, formatting.language)
				: run.equation
					? equation
						? ''
						: '[Equation]'
					: run.text);
		const script =
			formatting.verticalAlign === 'superscript' ||
			(noteText !== undefined && formatting.verticalAlign === undefined)
				? 'super'
				: formatting.verticalAlign === 'subscript'
					? 'sub'
					: undefined;
		const image = run.image;
		// A text box is drawn through its frame; one whose geometry cannot be drawn is not.
		if (image?.unsupported === 'Text box' && !image.shape) reportOnce(TEXT_BOX_NOTE);
		const object =
			equation ??
			(run.noteSeparator
				? {
						partName: '',
						contentType: '',
						widthPx: 192,
						heightPx: 1,
						noteSeparator: run.noteSeparator,
					}
				: // A picture watermark floats behind the page text and takes no room in the header.
					image && !image.anchored && !image.watermark && !image.pictureWatermark
					? {
							partName: image.partName,
							...(image.pictureFill ? { pictureFill: image.pictureFill } : {}),
							...(image.altText ? { altText: image.altText } : {}),
							...(image.transform ? { transform: image.transform } : {}),
							...(() => {
								const shadow = drawingOuterShadow(image.effects, theme);
								return shadow ? { shadow } : {};
							})(),
							contentType: image.contentType,
							widthPx: image.widthPx,
							heightPx: image.heightPx,
							...(image.chart
								? {
										chart: docxChartPaint(
											image.chart,
											image.widthPx,
											image.heightPx,
											theme,
											image.altText,
										),
									}
								: {}),
							...(image.shape ? shapePaint(image, theme, textBoxes) : {}),
							...inkPaint(image),
						}
					: undefined);
		return {
			sourceLength: reviewSourceLength(run),
			// Hidden text (`w:vanish`) takes no space when printed, as in Word by default.
			text:
				image || formatting.vanish
					? ''
					: formatting.caps
						? allCapsText(text, formatting.language)
						: formatting.smallCaps
							? smallCapsText(text, formatting.language)
							: text,
			...((name === 'MERGEFIELD' || name === 'ADDRESSBLOCK' || name === 'GREETINGLINE') && {
				mergeField: true,
			}),
			...(run.noteReference?.kind === 'footnote'
				? { footnoteReferenceId: run.noteReference.id }
				: {}),
			...(object ? { object } : {}),
			...definedProps({
				bold: formatting.bold,
				italic: formatting.italic,
				fontFamily: family,
				fontSizePt: formatting.fontSize,
				textScalePercent: scale === 100 ? undefined : scale,
				ligatures: formatting.ligatures,
				textShadow: formatting.textShadow,
				smallCaps: formatting.caps ? undefined : formatting.smallCaps,
				background,
				characterSpacingPx:
					formatting.characterSpacingTwips === undefined
						? undefined
						: formatting.characterSpacingTwips / 15,
				kerningThresholdPt: (formatting.kerningHalfPoints ?? 0) / 2,
				positionPx:
					formatting.positionHalfPoints === undefined
						? undefined
						: (formatting.positionHalfPoints * 2) / 3,
			}),
			...(color && /^#[0-9a-f]{6}$/i.test(color) ? { color } : {}),
			...(run.positionalTab ? { positionalTab: run.positionalTab } : {}),
			...(formatting.underline ? { underline: true } : {}),
			...(formatting.strike || formatting.doubleStrike ? { strike: true } : {}),
			...(run.break ? { breakAfter: run.break } : {}),
			...(script ? { script } : {}),
			...definedProps({
				kinsokuLanguage: kinsokuLanguage(
					resolveEastAsiaLanguage(run, {
						runCatalog,
						paragraphCatalog: catalog,
						paragraphStyleId,
					}),
				),
			}),
		};
	}
	const labels = computeListLabels(model);
	/** The list level's formatting overrides the paragraph's text or mark. */
	function labelRun(label: ParagraphListLabel, base: LayoutRun): LayoutRun {
		const marker = label.markerFormat;
		const suffix = label.suffix === 'tab' ? '\t' : label.suffix === 'space' ? ' ' : '';
		const markerFont = marker?.fontFamily ?? base.fontFamily;
		const pictureWidthPx = label.picture
			? (pictureBulletSizePt(
					label.picture,
					marker?.fontSizeHalfPoints
						? marker.fontSizeHalfPoints / 2
						: (base.fontSizePt ?? DEFAULT_FONT_SIZE_PT),
				).widthPt *
					4) /
				3
			: undefined;
		return {
			text: `${displayListLabel(label.text, markerFont)}${suffix}`,
			synthetic: true,
			marker: {
				length: displayListLabel(label.text, markerFont).length,
				alignment: label.alignment ?? 'left',
				...(pictureWidthPx !== undefined ? { widthPx: pictureWidthPx } : {}),
				...(label.legacy && label.suffix === 'tab'
					? {
							legacy: {
								indentPx: twipsToPx(label.legacy.indentTwips ?? NO_TWIPS),
								spacePx: twipsToPx(label.legacy.spaceTwips ?? NO_TWIPS),
							},
						}
					: {}),
			},
			...definedProps({
				bold: marker?.bold ?? base.bold,
				italic: marker?.italic ?? base.italic,
				fontFamily: markerFont,
				fontSizePt: marker?.fontSizeHalfPoints ? marker.fontSizeHalfPoints / 2 : base.fontSizePt,
			}),
			...(marker?.color || base.color ? { color: marker?.color ?? base.color! } : {}),
		};
	}
	const betweenBorders = new Map<LayoutParagraph, LayoutParagraphBorders['top']>();
	function withFontFallback(run: LayoutRun): LayoutRun {
		const fallback = run.fontFamily && model.fontFallbacks?.get(run.fontFamily.toLowerCase());
		return fallback ? { ...run, fontFallback: fallback } : run;
	}
	function adaptParagraph(
		paragraph: Paragraph,
		markLabel?: string,
		tableStyle?: TableStyleConditionalFormatting,
	): LayoutParagraph {
		const resolved =
			catalog || tableStyle?.paragraph
				? resolveParagraphFormatting(
						paragraph,
						catalog,
						tableStyle?.paragraph,
						model.overrideTableStyleFontSizeAndJustification,
					)
				: paragraph;
		const tableStyleRun = tableStyle?.run;
		const label = labels.get(paragraph.id);
		// Numbering level indents apply unless the paragraph or its style sets its own. A paragraph
		// numbered directly (`w:numPr`) takes the level's indents over its style's (Word draws the
		// text of a List Paragraph at the level's indent, not the style's 720 twips): only the
		// paragraph's own indents beat the level.
		const levelWins = Boolean(label && paragraph.numbering);
		const indents = levelWins ? paragraph : resolved;
		const ownFirstLine = indents.firstLineTwips !== undefined || indents.hangingTwips !== undefined;
		// A symbol that shares its run with text is drawn in its own font, the text in the run's.
		const runs = adaptContentControlTags(paragraph.runs, (run) =>
			symbolSegments(run).flatMap((segment) => {
				const adapted = adaptRun(segment, paragraph.style, markLabel, tableStyleRun);
				// Fields, note markers and symbol codes have their own display/source rules.
				if (segment.field || segment.noteReference || segment.noteMark || symbolOnly(segment))
					return [adapted];
				return scriptLayoutRuns(
					adapted,
					effective(segment, paragraph.style, tableStyleRun).formatting,
					theme,
				);
			}),
		);
		const visible = visibleReviewParagraph(paragraph, reviewMode);
		const footnotes = paragraphFootnotes(visible, visibleModel, noteLabel, adaptParagraph);
		const { betweenBorder, ...box } = paragraphBox(resolved, model.theme);
		const edges = blockTags.get(paragraph);
		const taggedRuns = edges ? adaptBlockContentControlTags(runs, edges) : runs;
		const mark = runs.some((run) => !run.hidden && (run.text || run.object))
			? undefined
			: adaptRun(paragraphMarkRun(paragraph), paragraph.style, undefined, tableStyleRun);
		const marker = label
			? labelRun(
					label,
					mark ??
						adaptRun(paragraph.runs[0] ?? { text: '' }, paragraph.style, undefined, tableStyleRun),
				)
			: undefined;
		const adapted: LayoutParagraph = {
			kind: 'paragraph',
			id: paragraph.id,
			...(!squeeze ? { squeezeJustifiedSpaces: false as const } : {}),
			// Measured in Word 16: distributed paragraphs still expand lines ending in a soft break.
			...(compatibility.doNotExpandShiftReturn && resolved.justification !== 'distribute'
				? { doNotExpandShiftReturn: true }
				: {}),
			runs: (marker ? [marker, ...taggedRuns] : taggedRuns).map(withFontFallback),
			...(footnotes.length ? { footnotes } : {}),
			...box,
			...(paragraph.tabStops?.length || label?.tabStopTwips !== undefined
				? {
						tabStops: [
							...(paragraph.tabStops ?? []).map((stop) => ({
								posPx: twipsToPx(stop.posTwips),
								align: stop.align,
								...(stop.leader ? { leader: stop.leader } : {}),
							})),
							...(label?.tabStopTwips !== undefined
								? [{ posPx: twipsToPx(label.tabStopTwips), align: 'left' as const }]
								: []),
						].sort((a, b) => a.posPx - b.posPx),
					}
				: {}),
			// Every run counts in the anchors' source offsets, as in the lines; hidden ones float nothing.
			...floatsOf(paragraph, model.theme, textBoxes, honoursLayoutInCell(compatibility), (run) =>
				isRunHiddenForReview(run, reviewMode),
			),
			...definedProps({
				defaultTabStopPx: model.defaultTabStopTwips
					? twipsToPx(model.defaultTabStopTwips)
					: undefined,
				emptyLineFont: !marker && mark ? fontOf(withFontFallback(mark)) : undefined,
				align: resolved.align ?? (isDisplayEquationParagraph(paragraph) ? 'center' : undefined),
				direction: resolved.direction,
				spacingBeforeTwips: autoSpaced(resolved.spacingBeforeAuto, resolved.spacingBeforeTwips),
				spacingAfterTwips: autoSpaced(resolved.spacingAfterAuto, resolved.spacingAfterTwips),
				lineSpacingTwips: resolved.lineSpacingTwips,
				lineSpacingRule: resolved.lineSpacingRule,
				indentLeftTwips:
					indents.indentLeftTwips ??
					label?.indentLeftTwips ??
					(levelWins ? resolved.indentLeftTwips : undefined),
				indentRightTwips: resolved.indentRightTwips,
				indentStartTwips: resolved.indentStartTwips,
				indentEndTwips: resolved.indentEndTwips,
				firstLineTwips: ownFirstLine ? indents.firstLineTwips : label?.firstLineTwips,
				hangingTwips: ownFirstLine ? indents.hangingTwips : label?.hangingTwips,
				styleId: contextualStyleId(paragraph, catalog),
			}),
			...(resolved.pageBreakBefore ? { pageBreakBefore: true } : {}),
			...(resolved.keepNext ? { keepNext: true } : {}),
			...(resolved.keepLines ? { keepLines: true } : {}),
			...(resolved.widowControl === false ? { widowControl: false } : {}),
			...(resolved.contextualSpacing ? { contextualSpacing: true } : {}),
			...(resolved.kinsoku === false ? { kinsoku: false as const } : {}),
			...(resolved.autoSpaceDE === false ? { autoSpaceDE: false as const } : {}),
			...(resolved.autoSpaceDN === false ? { autoSpaceDN: false as const } : {}),
			...(resolved.overflowPunct === false ? { overflowPunct: false as const } : {}),
			...(model.kinsoku ? { kinsokuSettings: model.kinsoku } : {}),
			...(model.autoHyphenation && !resolved.suppressAutoHyphens
				? {
						hyphenation: {
							...(model.hyphenationZoneTwips !== undefined
								? { zoneTwips: model.hyphenationZoneTwips }
								: {}),
							...(model.consecutiveHyphenLimit !== undefined
								? { consecutiveLimit: model.consecutiveHyphenLimit }
								: {}),
							...(model.doNotHyphenateCaps ? { doNotHyphenateCaps: true } : {}),
						},
					}
				: {}),
			...(!paragraph.dropCap && resolved.frame
				? { textFrame: adaptTextFrame(resolved.frame) }
				: {}),
		};
		if (betweenBorder) betweenBorders.set(adapted, betweenBorder);
		return adapted;
	}
	function adaptBlock(block: Block): LayoutBlock {
		if (block.type === 'paragraph') return adaptParagraph(block);
		reportOnce(TABLE_ROW_NOTE);
		reportOnce(VERTICAL_MERGE_NOTE);
		if (block.rows.some((row) => row.some((cell) => cell.nestedTables?.length)))
			reportOnce(
				'Nested tables are projected as read-only cell paragraphs in Print Layout; vertical merge and repeated header geometry are not rendered.',
			);
		return adaptTable(block, model, adaptParagraph);
	}

	const blocks = [
		...model.blocks.map(adaptBlock),
		...endnoteParagraphs(visibleModel, noteLabel, adaptParagraph),
	];
	groupParagraphBorders(blocks, betweenBorders);
	const folded = foldDropCaps(model.blocks, blocks);
	if (folded.size) reportOnce(DROP_CAP_NOTE);
	if (
		Object.values(model.footnoteSeparators ?? {}).some((note) =>
			note.blocks.some((block) => block.type === 'table'),
		)
	)
		reportOnce(
			'Tables in custom footnote separators are preserved but are not rendered in Print Layout.',
		);
	const footnoteSeparators = model.footnoteSeparators
		? Object.fromEntries(
				Object.entries(model.footnoteSeparators).map(([kind, note]) => [
					kind,
					note.blocks
						.filter((block): block is Paragraph => block.type === 'paragraph')
						.map((paragraph) => adaptParagraph(paragraph)),
				]),
			)
		: undefined;
	const sections = model.sections;
	const modernFootnoteSeparator = usesModernFootnoteLayout(compatibility);
	const legacyFootnoteColumns = attachesFootnotesToBodyColumns(compatibility);
	if (
		!modernFootnoteSeparator &&
		!legacyFootnoteColumns &&
		model.sections?.some((section) => section.footnoteColumns !== undefined)
	)
		reportOnce(
			'Footnote column settings are preserved, but this compatibility mode has not been verified against native Word; Print Layout uses the existing page-wide fallback.',
		);
	if (!sections || !sections.length) {
		return {
			sections: [
				{
					page: {
						widthPx: model.page.width,
						heightPx: model.page.height,
						marginTopPx: model.page.marginTop,
						marginRightPx: model.page.marginRight,
						marginBottomPx: model.page.marginBottom,
						marginLeftPx: model.page.marginLeft,
					},
					...(footnoteSeparators ? { footnoteSeparators } : {}),
					...(addParagraphSpacing ? { addParagraphSpacing: true } : {}),
					...(keepSpaceBeforePageBreak ? { keepSpaceBeforePageBreak: true } : {}),
					...(keepWrappingFloatsAboveBottomMargin
						? { keepWrappingFloatsAboveBottomMargin: true }
						: {}),
					...(modernFootnoteSeparator ? { modernFootnoteSeparator: true } : {}),
					...(legacyFootnoteColumns ? { legacyFootnoteColumns: true } : {}),
					blocks: blocks.filter((block) => !folded.has(block)),
				},
			],
		};
	}
	const indexById = new Map(model.blocks.map((block, index) => [block.id, index]));
	let start = 0;
	const result: LayoutSection[] = sections.map((section, index) => {
		const nextSection = sections[index + 1];
		const endIndex = nextSection
			? (indexById.get(section.endsAtBlockId) ?? start - 1) + 1
			: blocks.length;
		const end = Math.max(start, endIndex);
		const slice = blocks.slice(start, end);
		start = end;
		const lastBlock = model.blocks[end - 1];
		const markerId =
			nextSection &&
			lastBlock?.id === section.endsAtBlockId &&
			lastBlock.type === 'paragraph' &&
			isEmptyParagraph(lastBlock)
				? lastBlock.id
				: undefined;
		if (legacyFootnoteColumns && section.footnoteColumns)
			reportOnce(
				'Independent footnote column counts are preserved but ignored in measured compatibility modes 12 and 14; Print Layout attaches notes to body columns.',
			);
		const footnoteWidth =
			twipsToPx(section.pageWidthTwips) -
			twipsToPx(section.marginLeftTwips) -
			twipsToPx(section.marginRightTwips);
		if (
			modernFootnoteSeparator &&
			section.footnoteColumns &&
			(section.footnoteColumns > 45 || footnoteWidth <= 48 * (section.footnoteColumns - 1))
		)
			reportOnce(
				'Footnote columns exceeding the available page width or 45 rendered columns use one column in Print Layout; the saved count is preserved.',
			);
		const sectionStart = index === 0 ? undefined : sectionBreak(section.type, reportOnce);
		return {
			...(footnoteSeparators ? { footnoteSeparators } : {}),
			...(modernFootnoteSeparator ? { modernFootnoteSeparator: true } : {}),
			...(legacyFootnoteColumns ? { legacyFootnoteColumns: true } : {}),
			...(section.footnoteColumns !== undefined && modernFootnoteSeparator
				? {
						footnoteColumns: {
							count:
								section.footnoteColumns &&
								(section.footnoteColumns > 45 ||
									footnoteWidth <= 48 * (section.footnoteColumns - 1))
									? 1
									: Math.max(1, section.footnoteColumns || section.columns.count),
							gapPx: section.footnoteColumns
								? 48
								: twipsToPx(section.columns.spacingTwips ?? twips(720)),
							...(!section.footnoteColumns &&
							!section.columns.equalWidth &&
							section.columns.widths?.length
								? {
										widths: section.columns.widths.map((column) => ({
											widthPx: twipsToPx(column.widthTwips),
											gapPx: twipsToPx(
												column.spacingTwips ?? section.columns.spacingTwips ?? twips(720),
											),
										})),
									}
								: {}),
						},
					}
				: {}),
			page: {
				widthPx: twipsToPx(section.pageWidthTwips),
				heightPx: twipsToPx(section.pageHeightTwips),
				marginTopPx: twipsToPx(section.marginTopTwips),
				marginRightPx: twipsToPx(section.marginRightTwips),
				marginBottomPx: twipsToPx(section.marginBottomTwips),
				marginLeftPx: twipsToPx(section.marginLeftTwips),
			},
			...(section.columns.count > 1
				? {
						columns: {
							count: section.columns.count,
							gapPx: twipsToPx(section.columns.spacingTwips ?? twips(720)),
							...(section.columns.separator ? { separator: true } : {}),
							...(!section.columns.equalWidth && section.columns.widths?.length
								? {
										widths: section.columns.widths.map((column) => ({
											widthPx: twipsToPx(column.widthTwips),
											gapPx: twipsToPx(
												column.spacingTwips ?? section.columns.spacingTwips ?? twips(720),
											),
										})),
									}
								: {}),
						},
					}
				: {}),
			...(addParagraphSpacing ? { addParagraphSpacing: true } : {}),
			...(keepSpaceBeforePageBreak ? { keepSpaceBeforePageBreak: true } : {}),
			...(keepWrappingFloatsAboveBottomMargin ? { keepWrappingFloatsAboveBottomMargin: true } : {}),
			...(sectionStart && { break: sectionStart }),
			...(section.verticalAlign ? { verticalAlign: section.verticalAlign } : {}),
			// Empty text break markers share the preceding line. The mandatory
			// paragraph after a table has its own line, after balancing the table.
			blocks: slice.flatMap((block, offset) => {
				if (folded.has(block)) return [];
				if (block.kind !== 'paragraph' || block.id !== markerId) return [block];
				if (slice[offset - 1]?.kind !== 'table') return [];
				return nextSection?.type === 'continuous'
					? [{ ...block, afterTableSectionBreak: true }]
					: [block];
			}),
		};
	});
	// Word collapses the first paragraph's space before with the previous section's last space after.
	return {
		sections: result.map((section, index) => {
			const last = result[index - 1]?.blocks.at(-1);
			const after = last?.kind === 'paragraph' ? last.spacingAfterTwips : undefined;
			return !addParagraphSpacing && after
				? { ...section, leadingSpacingAfterPx: twipsToPx(after) }
				: section;
		}),
	};
}
