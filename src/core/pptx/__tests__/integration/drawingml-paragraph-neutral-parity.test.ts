import { beforeAll, describe, expect, it } from 'vitest';

import { parseDrawingParagraphProperties } from '../../../drawingml/read-paragraph-properties';
import type { DrawingParagraphProperties } from '../../../drawingml/text-model';
import { NS, parseXml } from '../../../xml/index';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import {
	applyBulletProperties,
	buildParagraphPropertiesXml,
} from '../../core/core/runtime/PptxHandlerRuntimeSaveParagraphHelpers';
import type { BulletInfo, TextStyle, XmlObject } from '../../core/types';
import {
	bulletFieldsFromDrawing,
	parseAlignmentAttr,
	parseBulletInfo,
	parseLineSpacingExactPt,
	parseLineSpacingMultiplier,
	parseParagraphExtraAttributes,
	parseParagraphLevel,
	parseParagraphMargins,
	parseParagraphRtl,
	parseTabStops,
} from '../../core/utils/paragraph-properties-parser';
import {
	objectParagraphProperties,
	objectTextSpacing,
} from '../../core/utils/paragraph-properties-neutral';
import { canonicalJson, walkParityCorpus, type ParityVisit } from './drawingml-text-parity-corpus';

/**
 * DrawingML paragraph property parity (docs/agnostic-core-plan.md step 3): the pptx readers and
 * writers of the paragraph attributes, spacing, the bullet group (size, font, type and the
 * follow-text markers) and tab stops run on the shared codec (`drawingml/text-paragraph-codec`),
 * and this harness proves the two paths agree over every `a:pPr`, `a:defPPr` and `a:lvlNpPr` of
 * every part of every committed deck (`pptx/__tests__/fixtures`, including the `e2e` snapshot).
 *
 * - Reading: each element is read twice, through the pptx object tree
 *   (`objectParagraphProperties`, what the pptx parsers run) and through the DOM
 *   (`parseDrawingParagraphProperties`), paired by element path, and the codec families must be
 *   deeply equal.
 * - Saving: the pptx model read from the object tree (the `TextStyle` fields, `BulletInfo` and tab
 *   stops) is written back through the pptx paragraph writers (`buildParagraphPropertiesXml`,
 *   `applyBulletProperties`), serialised, and read through the DOM reader; the codec families
 *   must equal the source's, except the ones the pptx paragraph writer does not write from these
 *   fields (`a:spcBef`/`a:spcAft` percentages, saved from the shape style; out-of-range line
 *   spacing, clamped by PowerPoint's [0.1, 5] rule, which no committed deck has).
 *
 * Today: 6,647 parts, 54,742 elements and 191,150 codec fields read, 54,742 saves; no
 * differences. No committed deck repeats a bullet type, size or font choice or writes both
 * `a:spcPts` and `a:spcPct` in one spacing element, so the readers' choice precedence (schema
 * order over the object tree, document order over the DOM) is never exercised.
 *
 * Before the switch, the previous pptx parsers and writers were compared with the new ones over the
 * same corpus (54,742 paragraph property elements): every read value and every written value was
 * equal. Only the attribute order of a rebuilt `a:pPr` changed (11,975 elements), to schema order;
 * 10,917 of those now match the source file's own order, which the old order never did.
 */

const PARAGRAPH_ELEMENTS = new Set([
	'pPr',
	'defPPr',
	...Array.from({ length: 9 }, (_, index) => `lvl${index + 1}pPr`),
]);
const builder = new PptxRuntimeDependencyFactory().createBuilder();
const canonical = canonicalJson;

/** The codec families of a DOM reading (no colour, picture, default run or kept children). */
function codecFamilies(value: DrawingParagraphProperties | undefined): DrawingParagraphProperties {
	const {
		extraXml: _extra,
		defaultRunProperties: _run,
		bulletColor,
		bullet,
		...rest
	} = value ?? {};
	return {
		...rest,
		...(bulletColor === 'text' ? { bulletColor } : {}),
		...(bullet && bullet.kind !== 'picture' ? { bullet } : {}),
	};
}

/** What the pptx paragraph writer saves of a paragraph properties node, read back on the DOM. */
function savedThroughPptx(node: XmlObject): DrawingParagraphProperties {
	const style: TextStyle = {
		...parseParagraphMargins(node),
		...parseParagraphExtraAttributes(node),
		rtl: parseParagraphRtl(node),
		tabStops: parseTabStops(node),
		lineSpacing: parseLineSpacingMultiplier(node['a:lnSpc'] as XmlObject | undefined),
		lineSpacingExactPt: parseLineSpacingExactPt(node['a:lnSpc'] as XmlObject | undefined),
	};
	const level = node['@_lvl'] !== undefined ? parseParagraphLevel(node) : undefined;
	const align = objectParagraphProperties(node).align;
	if (align) style.align = parseAlignmentAttr(align) ?? 'left';
	const pPr = buildParagraphPropertiesXml(
		style,
		align,
		undefined,
		{
			spacingBefore: undefined,
			spacingAfter: undefined,
			lineSpacing:
				style.lineSpacing === undefined
					? undefined
					: objectTextSpacing({ unit: 'percent', value: style.lineSpacing }),
			lineSpacingExactPt: style.lineSpacingExactPt,
		},
		level,
		style,
	);
	for (const key of ['a:spcBef', 'a:spcAft'] as const) {
		if ((node[key] as XmlObject | undefined)?.['a:spcPts']) pPr[key] = node[key];
	}
	const info = parseBulletInfo(node, 0);
	const neutral = objectParagraphProperties(node);
	if (info || neutral.bulletSize || neutral.bulletFont || neutral.bulletColor)
		applyBulletProperties(pPr, {
			...(info ?? {}),
			...bulletFieldsFromDrawing(objectParagraphProperties(node)),
		} as BulletInfo);
	if (style.tabStops === undefined && node['a:tabLst'] !== undefined) pPr['a:tabLst'] = {};
	const xml = builder.build({ 'a:pPr': { '@_xmlns:a': NS.a, ...pPr } }) as string;
	return codecFamilies(parseDrawingParagraphProperties(parseXml(xml).documentElement));
}

/** The source families the pptx paragraph writer saves from the fields above. */
function savedFamilies(value: DrawingParagraphProperties): DrawingParagraphProperties {
	const out = { ...value };
	if (out.spaceBefore?.unit === 'percent') delete out.spaceBefore;
	if (out.spaceAfter?.unit === 'percent') delete out.spaceAfter;
	if (out.tabStops) out.tabStops = out.tabStops.filter((stop) => stop.positionEmu !== undefined);
	return out;
}

interface Difference {
	label: string;
	objectTree: string;
	dom: string;
}

const readDifferences: Difference[] = [];
const saveDifferences: Difference[] = [];
let elementsCompared = 0;
let fieldsCompared = 0;
let savesCompared = 0;
let partsCompared = 0;

function compareParagraph({ label, element, node }: ParityVisit): void {
	const dom = codecFamilies(parseDrawingParagraphProperties(element));
	const objectTree = objectParagraphProperties(node);
	elementsCompared++;
	fieldsCompared += Object.keys(dom).length;
	if (canonical(objectTree) !== canonical(dom))
		readDifferences.push({ label, objectTree: canonical(objectTree), dom: canonical(dom) });
	const lineSpacing = dom.lineSpacing;
	const clamped =
		lineSpacing?.unit === 'percent' && (lineSpacing.value < 0.1 || lineSpacing.value > 5);
	if (clamped) return;
	savesCompared++;
	const saved = savedThroughPptx(node);
	if (canonical(saved) !== canonical(savedFamilies(dom)))
		saveDifferences.push({
			label,
			objectTree: canonical(saved),
			dom: canonical(savedFamilies(dom)),
		});
}

beforeAll(async () => {
	partsCompared = await walkParityCorpus(/pPr\b/u, PARAGRAPH_ELEMENTS, compareParagraph);
}, 600_000);

describe('DrawingML paragraph properties: pptx object tree vs shared DOM reader', () => {
	it('compares a real corpus', () => {
		expect(partsCompared).toBeGreaterThan(6000);
		expect(elementsCompared).toBeGreaterThan(54_000);
		expect(fieldsCompared).toBeGreaterThan(190_000);
		expect(savesCompared).toBeGreaterThan(54_000);
	});

	it('reads every paragraph properties element the same way', () => {
		expect(readDifferences.slice(0, 5)).toEqual([]);
		expect(readDifferences).toHaveLength(0);
	});

	it('saves every paragraph properties element back to the same codec values', () => {
		expect(saveDifferences.slice(0, 5)).toEqual([]);
		expect(saveDifferences).toHaveLength(0);
	});
});
