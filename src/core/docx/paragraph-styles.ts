import { readStyleMetadata } from './style-metadata';
import { parsePercentageShading } from './percentage-shading';
import type {
	Paragraph,
	ParagraphFormatting,
	ParagraphFrame,
	ParagraphStyleCatalog,
	ParagraphStyleDefinition,
} from './model';
import {
	isStHAnchor,
	isStHeightRule,
	isStVAnchor,
	isStWrap,
	isStXAlign,
	isStYAlign,
} from './generated/wml-simple-types';
import { children, first, getW, parseXml, type XmlElement } from './xml';
import {
	cachedResolver,
	chainResolver,
	defaultStyleId,
	inheritanceCycleEntries,
} from './style-inheritance';
import { parseParagraphBorders, parseShadingFill } from './table-borders';
import { alignFromJustification, parseJustification } from './paragraph-alignment';
import {
	onOffElement,
	parseInteger,
	parseOnOff,
	parseSignedTwips,
	parseTwips,
} from './simple-types';

const signedTwipValue = parseSignedTwips;
const twipValue = parseTwips;
const enabled = onOffElement;

/** Pagination toggles whose element names match their model keys. */
export const PAGINATION_KEYS = [
	'keepNext',
	'keepLines',
	'widowControl',
	'contextualSpacing',
	'suppressLineNumbers',
] as const;

/** East Asian typography toggles (`w:kinsoku`...); read-only, kept in the source XML on save. */
export const EAST_ASIAN_TYPOGRAPHY_KEYS = [
	'kinsoku',
	'autoSpaceDE',
	'autoSpaceDN',
	'overflowPunct',
] as const;

/** `w:outlineLvl` as a 1-9 level (0 for body text, `w:val="9"`); undefined when absent or invalid. */
export function parseOutlineLevel(element: XmlElement | undefined): number | undefined {
	const raw = getW(element, 'val');
	if (raw === undefined) return undefined;
	const value = Number(raw);
	if (!Number.isInteger(value) || value < 0 || value > 9) return undefined;
	return value === 9 ? 0 : value + 1;
}

/** Non-drop-cap `w:framePr` text frame properties; returns undefined for drop caps or empty elements. */
export function parseFrameProperties(element: XmlElement | undefined): ParagraphFrame | undefined {
	if (!element) return undefined;
	const dropCap = getW(element, 'dropCap');
	if (dropCap === 'drop' || dropCap === 'margin') return undefined;
	const frame: ParagraphFrame = {};
	const widthTwips = twipValue(getW(element, 'w'));
	const heightTwips = twipValue(getW(element, 'h'));
	const heightRule = getW(element, 'hRule');
	const xTwips = signedTwipValue(getW(element, 'x'));
	const yTwips = signedTwipValue(getW(element, 'y'));
	const xAlign = getW(element, 'xAlign');
	const yAlign = getW(element, 'yAlign');
	const hAnchor = getW(element, 'hAnchor');
	const vAnchor = getW(element, 'vAnchor');
	const hSpaceTwips = twipValue(getW(element, 'hSpace'));
	const vSpaceTwips = twipValue(getW(element, 'vSpace'));
	const wrap = getW(element, 'wrap');
	if (widthTwips !== undefined) frame.widthTwips = widthTwips;
	if (heightTwips !== undefined) frame.heightTwips = heightTwips;
	if (isStHeightRule(heightRule)) frame.heightRule = heightRule;
	if (xTwips !== undefined) frame.xTwips = xTwips;
	if (yTwips !== undefined) frame.yTwips = yTwips;
	if (isStXAlign(xAlign)) frame.xAlign = xAlign;
	if (isStYAlign(yAlign)) frame.yAlign = yAlign;
	if (isStHAnchor(hAnchor)) frame.hAnchor = hAnchor;
	if (isStVAnchor(vAnchor)) frame.vAnchor = vAnchor;
	if (hSpaceTwips !== undefined) frame.hSpaceTwips = hSpaceTwips;
	if (vSpaceTwips !== undefined) frame.vSpaceTwips = vSpaceTwips;
	if (isStWrap(wrap)) frame.wrap = wrap;
	return Object.keys(frame).length > 0 ? frame : undefined;
}

/** Parses `w:pPr` formatting (shared by styles, docDefaults and direct paragraph properties). */
export function parseFormatting(pPr: XmlElement | undefined): ParagraphFormatting {
	const result: ParagraphFormatting = {};
	const { align, justification } = parseJustification(pPr);
	if (justification) result.justification = justification;
	if (align) result.align = align;
	const bidi = enabled(first(pPr, 'bidi'));
	if (bidi !== undefined) result.direction = bidi ? 'rtl' : 'ltr';
	const pageBreakBefore = enabled(first(pPr, 'pageBreakBefore'));
	if (pageBreakBefore !== undefined) result.pageBreakBefore = pageBreakBefore;
	for (const key of [...PAGINATION_KEYS, ...EAST_ASIAN_TYPOGRAPHY_KEYS]) {
		const value = enabled(first(pPr, key));
		if (value !== undefined) result[key] = value;
	}
	const suppressAutoHyphens = enabled(first(pPr, 'suppressAutoHyphens'));
	if (suppressAutoHyphens !== undefined) result.suppressAutoHyphens = suppressAutoHyphens;
	const outline = parseOutlineLevel(first(pPr, 'outlineLvl'));
	if (outline !== undefined) result.outlineLevel = outline;
	const frame = parseFrameProperties(first(pPr, 'framePr'));
	if (frame) result.frame = frame;
	const borders = parseParagraphBorders(first(pPr, 'pBdr'));
	if (borders) result.borders = borders;
	const shadingElement = first(pPr, 'shd');
	const shading = parseShadingFill(shadingElement);
	const pattern = parsePercentageShading(shadingElement);
	if (pattern) result.percentageShading = pattern;
	if (shading) result.shadingFill = shading;
	const spacing = first(pPr, 'spacing');
	const before = twipValue(getW(spacing, 'before'));
	const after = twipValue(getW(spacing, 'after'));
	const line = signedTwipValue(getW(spacing, 'line'));
	if (before !== undefined) result.spacingBeforeTwips = before;
	if (after !== undefined) result.spacingAfterTwips = after;
	const beforeAuto = parseOnOff(getW(spacing, 'beforeAutospacing'));
	const afterAuto = parseOnOff(getW(spacing, 'afterAutospacing'));
	if (beforeAuto !== undefined) result.spacingBeforeAuto = beforeAuto;
	if (afterAuto !== undefined) result.spacingAfterAuto = afterAuto;
	if (line !== undefined) result.lineSpacingTwips = line;
	const rule = getW(spacing, 'lineRule');
	if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') result.lineSpacingRule = rule;
	else if (line !== undefined) result.lineSpacingRule = 'auto';
	const ind = first(pPr, 'ind');
	for (const [xml, key] of [
		['left', 'indentLeftTwips'],
		['right', 'indentRightTwips'],
		['start', 'indentStartTwips'],
		['end', 'indentEndTwips'],
	] as const) {
		const value = signedTwipValue(getW(ind, xml));
		if (value !== undefined) result[key] = value;
	}
	for (const [xml, key] of [
		['firstLine', 'firstLineTwips'],
		['hanging', 'hangingTwips'],
	] as const) {
		const value = twipValue(getW(ind, xml));
		if (value !== undefined) result[key] = value;
	}
	return result;
}

function parseStyleNumbering(
	pPr: XmlElement | undefined,
): { numId: number; level: number } | undefined {
	const numPr = first(pPr, 'numPr');
	if (!numPr) return undefined;
	const numId = parseInteger(getW(first(numPr, 'numId'), 'val'));
	if (numId === undefined) return undefined;
	return { numId, level: parseInteger(getW(first(numPr, 'ilvl'), 'val')) ?? 0 };
}

export function parseParagraphStyleCatalog(xml: string | XmlElement): ParagraphStyleCatalog {
	const root = typeof xml === 'string' ? parseXml(xml).documentElement : xml;
	const defaults = first(first(root, 'docDefaults'), 'pPrDefault');
	const pDefaults = first(defaults, 'pPr');
	const styles = Object.create(null) as Record<string, ParagraphStyleDefinition>;
	for (const element of children(root, 'style')) {
		if (getW(element, 'type') !== 'paragraph') continue;
		const id = getW(element, 'styleId');
		if (!id) continue;
		const next = getW(first(element, 'next'), 'val');
		const linkedStyle = getW(first(element, 'link'), 'val');
		const numbering = parseStyleNumbering(first(element, 'pPr'));
		styles[id] = {
			...readStyleMetadata(element, id),
			...(next !== undefined && { next }),
			...(linkedStyle !== undefined && { linkedStyle }),
			...(numbering ? { numbering } : {}),
			formatting: parseFormatting(first(element, 'pPr')),
		};
	}
	const warnings = inheritanceCycleEntries(styles).map(
		(id) => `Paragraph style inheritance cycle detected at “${id}”; cyclic inheritance is ignored.`,
	);
	return { docDefaults: parseFormatting(pDefaults), styles, warnings };
}

type Resolver<R> = (id: string | undefined) => R | undefined;
const formattingResolvers = new WeakMap<object, Resolver<ParagraphFormatting>>();
const numberingResolvers = new WeakMap<
	object,
	Resolver<Pick<ParagraphStyleDefinition, 'numbering'>>
>();

/** The formatting a style's chain gives, each style's `formatting` assigned root first. */
function styleFormatting(
	styleId: string | undefined,
	catalog: ParagraphStyleCatalog,
): ParagraphFormatting | undefined {
	return cachedResolver(formattingResolvers, catalog.styles, () =>
		chainResolver(
			catalog.styles,
			(style) => style.formatting,
			(resolved, formatting) => {
				const merged = Object.assign({}, resolved, formatting);
				if (resolved.frame && formatting.frame)
					merged.frame = { ...resolved.frame, ...formatting.frame };
				return merged;
			},
			{},
		),
	)(styleId);
}

/** Resolves defaults and style ancestry without modifying paragraph direct properties. */
export function resolveParagraphFormatting(
	paragraph: Paragraph,
	catalog: ParagraphStyleCatalog,
): ParagraphFormatting {
	const result: ParagraphFormatting = { ...catalog.docDefaults };
	const explicitStyle = paragraph.style;
	Object.assign(result, styleFormatting(explicitStyle || defaultStyleId(catalog.styles), catalog));
	const direct: ParagraphFormatting = {};
	const keys: (keyof ParagraphFormatting)[] = [
		'align',
		'justification',
		'direction',
		'spacingBeforeTwips',
		'spacingAfterTwips',
		'spacingBeforeAuto',
		'spacingAfterAuto',
		'lineSpacingTwips',
		'lineSpacingRule',
		'indentLeftTwips',
		'indentRightTwips',
		'indentStartTwips',
		'indentEndTwips',
		'firstLineTwips',
		'hangingTwips',
		'pageBreakBefore',
		...PAGINATION_KEYS,
		'suppressAutoHyphens',
		...EAST_ASIAN_TYPOGRAPHY_KEYS,
		'outlineLevel',
		'frame',
		'borders',
		'shadingFill',
		'percentageShading',
	];
	for (const key of keys) {
		const value = paragraph[key];
		if (value !== undefined) Object.assign(direct, { [key]: value });
	}
	const inheritedFrame = result.frame;
	Object.assign(result, direct);
	if (inheritedFrame && direct.frame) result.frame = { ...inheritedFrame, ...direct.frame };
	// A direct alignment without its own `w:jc` value supersedes any inherited exact value.
	if (paragraph.align !== undefined && paragraph.justification === undefined)
		delete result.justification;
	// `start`/`end` follow the resolved direction, which may be inherited from a style.
	if (
		(result.justification === 'start' || result.justification === 'end') &&
		(paragraph.align === undefined ||
			(paragraph.justification !== undefined &&
				paragraph.align ===
					alignFromJustification(paragraph.justification, paragraph.direction === 'rtl')))
	) {
		const align = alignFromJustification(result.justification, result.direction === 'rtl');
		if (align === undefined) delete result.align;
		else result.align = align;
	}
	return result;
}

/** Resolves numbering inherited through `pStyle` when a paragraph has no direct `w:numPr`. */
export function resolveStyleNumbering(
	styleId: string | undefined,
	catalog: ParagraphStyleCatalog,
): { numId: number; level: number } | undefined {
	return cachedResolver(numberingResolvers, catalog.styles, () =>
		chainResolver<ParagraphStyleDefinition, Pick<ParagraphStyleDefinition, 'numbering'>>(
			catalog.styles,
			(style) => (style.numbering ? { numbering: style.numbering } : undefined),
			(_resolved, nearer) => nearer,
			{},
		),
	)(styleId || undefined)?.numbering;
}
