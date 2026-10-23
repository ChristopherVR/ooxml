// Projections for the SmartArt text parity harness: one comparison shape built from the pptx
// object-tree text parsers and one from the neutral DrawingML text model, in pptx units (px).
import type {
	DrawingBodyProperties,
	DrawingParagraphProperties,
	DrawingRunProperties,
} from '../../../drawingml/index';
import type { XmlObject } from '../../core/types';
import { parseBodyPrBooleanAttrs } from '../../core/utils/body-properties-parser';
import {
	parseAlignmentAttr,
	parseBulletInfo,
	parseLineSpacingExactPt,
	parseLineSpacingMultiplier,
	parseParagraphExtraAttributes,
	parseParagraphLevel,
	parseParagraphMargins,
	parseParagraphRtl,
	parseParagraphSpacingPx,
} from '../../core/utils/paragraph-properties-parser';
import {
	parseRunFontElements,
	parseRunHyperlink,
	parseRunPropertyAttributes,
	parseRunSolidFillColor,
	parseRunSymbolFont,
} from '../../core/utils/text-run-properties-parser';

const EMU_PER_PX = 9525;
const PX_PER_PT = 96 / 72;

type Shape = Record<string, unknown>;

/** Drops undefined entries and rounds numbers so both sides compare exactly. */
export function clean(shape: Shape): Shape {
	return Object.fromEntries(
		Object.entries(shape)
			.filter(([, value]) => value !== undefined)
			.map(([key, value]) => [
				key,
				typeof value === 'number' ? Math.round(value * 1e6) / 1e6 : value,
			]),
	);
}

function child(node: XmlObject | undefined, name: string): XmlObject | undefined {
	const key = node && Object.keys(node).find((candidate) => candidate.split(':').pop() === name);
	const value = key ? node[key] : undefined;
	if (key && value === '') return {};
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as XmlObject)
		: undefined;
}

// ---------------------------------------------------------------------------
// pptx side
// ---------------------------------------------------------------------------

export function pptxRun(rPr: XmlObject | undefined): Shape {
	const attributes = parseRunPropertyAttributes(rPr);
	const fonts = parseRunFontElements(rPr);
	const link = parseRunHyperlink(rPr);
	return clean({
		fontSize: attributes.fontSize,
		bold: attributes.bold,
		italic: attributes.italic,
		underline: attributes.underline,
		underlineStyle: attributes.underlineStyle,
		strikethrough: attributes.strikethrough,
		strikeType: attributes.strikeType,
		textCaps: attributes.textCaps,
		baseline: attributes.baseline,
		characterSpacing: attributes.characterSpacing,
		kerning: attributes.kerning,
		language: attributes.language,
		dirty: attributes.dirty,
		spellingError: attributes.spellingError,
		noProof: attributes.noProof,
		fontFamily: fonts.fontFamily,
		eastAsiaFont: fonts.eastAsiaFont,
		complexScriptFont: fonts.complexScriptFont,
		latinFontPanose: fonts.latinFontPanose,
		latinFontPitchFamily: fonts.latinFontPitchFamily,
		latinFontCharset: fonts.latinFontCharset,
		symbolFont: parseRunSymbolFont(rPr),
		hyperlinkRId: link.hyperlinkRId,
		color: parseRunSolidFillColor(rPr),
	});
}

export function pptxParagraph(pPr: XmlObject | undefined): Shape {
	const bullet = parseBulletInfo(pPr) ?? undefined;
	return clean({
		align: parseAlignmentAttr(pPr?.['@_algn'] as string | undefined),
		level: parseParagraphLevel(pPr),
		rtl: parseParagraphRtl(pPr),
		...parseParagraphMargins(pPr),
		...parseParagraphExtraAttributes(pPr),
		spaceBefore: parseParagraphSpacingPx(child(pPr, 'spcBef')),
		spaceAfter: parseParagraphSpacingPx(child(pPr, 'spcAft')),
		lineSpacing: parseLineSpacingMultiplier(child(pPr, 'lnSpc')),
		lineSpacingExactPt: parseLineSpacingExactPt(child(pPr, 'lnSpc')),
		bullet: bullet ? JSON.stringify(clean(bullet as Shape)) : undefined,
	});
}

export function pptxBody(bodyPr: XmlObject | undefined): Shape {
	const style: Record<string, unknown> = {};
	if (bodyPr) parseBodyPrBooleanAttrs(bodyPr, style);
	return clean(style);
}

// ---------------------------------------------------------------------------
// neutral side, converted to the pptx units and vocabulary
// ---------------------------------------------------------------------------

const ALIGN: Record<string, string> = {
	l: 'left',
	ctr: 'center',
	r: 'right',
	just: 'justify',
	justify: 'justify',
	justLow: 'justLow',
	dist: 'dist',
	thaiDist: 'thaiDist',
};

function token(value: string | undefined): string {
	return (value ?? '').trim().toLowerCase();
}

export function neutralRun(run: DrawingRunProperties | undefined): Shape {
	if (!run) return {};
	const u = token(run.underline);
	const strike = token(run.strike);
	const caps = token(run.caps);
	const underline =
		run.underline === undefined ? undefined : !['', 'none', '0', 'false'].includes(u);
	const struck =
		run.strike === undefined ? undefined : !['', 'nostrike', 'none', '0', 'false'].includes(strike);
	const color =
		run.fill?.kind === 'solid' && run.fill.color.kind === 'srgb' ? run.fill.color.value : undefined;
	return clean({
		fontSize: run.sizePt === undefined ? undefined : run.sizePt * PX_PER_PT,
		bold: run.bold,
		italic: run.italic,
		underline,
		underlineStyle: underline ? run.underline : undefined,
		strikethrough: struck,
		strikeType: struck ? (strike === 'dblstrike' ? 'dblStrike' : 'sngStrike') : undefined,
		textCaps: caps === 'all' || caps === 'small' ? caps : undefined,
		baseline: run.baseline === undefined ? undefined : Math.round(run.baseline * 100000),
		characterSpacing: run.spacingPt === undefined ? undefined : Math.round(run.spacingPt * 100),
		kerning: run.kernPt === undefined ? undefined : Math.round(run.kernPt * 100),
		language: run.language,
		dirty: run.dirty,
		spellingError: run.spellingError,
		noProof: run.noProof,
		fontFamily: run.latin?.typeface,
		eastAsiaFont: run.eastAsian?.typeface,
		complexScriptFont: run.complexScript?.typeface,
		latinFontPanose: run.latin?.panose,
		latinFontPitchFamily: run.latin?.pitchFamily,
		latinFontCharset: run.latin?.charset,
		symbolFont: run.symbol?.typeface,
		hyperlinkRId: run.hyperlinkClick?.relId,
		color,
	});
}

function neutralBullet(properties: DrawingParagraphProperties): string | undefined {
	const bullet = properties.bullet;
	if (!bullet || bullet.kind === 'picture') return undefined;
	if (bullet.kind === 'none') return JSON.stringify({ none: true });
	const font = properties.bulletFont;
	const size = properties.bulletSize;
	const color = properties.bulletColor;
	const shared = {
		fontFamily: typeof font === 'object' ? font.typeface : undefined,
		sizePercent: typeof size === 'object' && size.unit === 'percent' ? size.value * 100 : undefined,
		sizePts: typeof size === 'object' && size.unit === 'points' ? size.value : undefined,
		color: typeof color === 'object' && color.kind === 'srgb' ? color.value : undefined,
	};
	const value =
		bullet.kind === 'char'
			? { char: bullet.char, ...shared }
			: {
					autoNumType: bullet.scheme,
					autoNumStartAt: bullet.startAt ?? 1,
					paragraphIndex: 0,
					...shared,
				};
	return JSON.stringify(clean(value));
}

function points(spacing: DrawingParagraphProperties['spaceBefore']): number | undefined {
	return spacing?.unit === 'points' ? spacing.value * PX_PER_PT : undefined;
}

export function neutralParagraph(properties: DrawingParagraphProperties | undefined): Shape {
	const p = properties ?? {};
	const line = p.lineSpacing;
	return clean({
		align: p.align ? ALIGN[p.align] : undefined,
		level: p.level ?? 0,
		rtl: p.rightToLeft,
		paragraphMarginLeft: p.marginLeftEmu === undefined ? undefined : p.marginLeftEmu / EMU_PER_PX,
		paragraphMarginRight:
			p.marginRightEmu === undefined ? undefined : p.marginRightEmu / EMU_PER_PX,
		paragraphIndent: p.indentEmu === undefined ? undefined : p.indentEmu / EMU_PER_PX,
		defaultTabSize:
			p.defaultTabSizeEmu === undefined ? undefined : p.defaultTabSizeEmu / EMU_PER_PX,
		eaLineBreak: p.eastAsianLineBreak,
		latinLineBreak: p.latinLineBreak,
		fontAlignment: p.fontAlign,
		hangingPunctuation: p.hangingPunctuation,
		spaceBefore: points(p.spaceBefore),
		spaceAfter: points(p.spaceAfter),
		lineSpacing: line?.unit === 'percent' ? Math.max(0.1, Math.min(5, line.value)) : undefined,
		lineSpacingExactPt: line?.unit === 'points' && line.value > 0 ? line.value : undefined,
		bullet: neutralBullet(p),
	});
}

export function neutralBody(body: DrawingBodyProperties | undefined): Shape {
	if (!body) return {};
	return clean({
		compatibleLineSpacing: body.compatibleLineSpacing,
		forceAntiAlias: body.forceAntiAlias,
		upright: body.upright,
		fromWordArt: body.fromWordArt,
		spaceFirstLastParagraph: body.spaceFirstLastParagraph,
		anchorCenter: body.anchorCenter,
		rtlColumns: body.rightToLeftColumns,
		textBodyRotation: body.rotation,
	});
}
