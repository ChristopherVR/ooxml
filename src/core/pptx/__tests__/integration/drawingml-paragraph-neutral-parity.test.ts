import JSZip from 'jszip';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import { parseDrawingParagraphProperties } from '../../../drawingml/read-paragraph-properties';
import type { DrawingParagraphProperties } from '../../../drawingml/text-model';
import { NS, elements, parseXml, type XmlElement } from '../../../xml/index';
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

const CORE = path.resolve(__dirname, '../../..');
const PARAGRAPH_ELEMENTS = new Set([
	'pPr',
	'defPPr',
	...Array.from({ length: 9 }, (_, index) => `lvl${index + 1}pPr`),
]);
const parser = new PptxRuntimeDependencyFactory().createParser();
const builder = new PptxRuntimeDependencyFactory().createBuilder();

function decks(): string[] {
	const found: string[] = [];
	const walk = (dir: string) => {
		for (const name of readdirSync(dir).sort()) {
			const full = path.join(dir, name);
			if (statSync(full).isDirectory()) walk(full);
			else if (/\.(pptx|pptm|potx|ppsx)$/i.test(name)) found.push(full);
		}
	};
	walk(path.join(CORE, 'pptx/__tests__/fixtures'));
	return found;
}

/** The object-tree child for a DOM child: same-name siblings are an array, a bare element `''`. */
function objectChild(node: unknown, name: string, index: number): unknown {
	if (!node || typeof node !== 'object') return undefined;
	const value = (node as XmlObject)[name];
	return Array.isArray(value) ? value[index] : index === 0 ? value : undefined;
}

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

const canonical = (value: unknown): string =>
	JSON.stringify(value, (_key, entry: unknown) =>
		entry && typeof entry === 'object' && !Array.isArray(entry)
			? Object.fromEntries(
					Object.entries(entry as Record<string, unknown>).sort(([a], [b]) => a.localeCompare(b)),
				)
			: entry,
	);

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

function compareParagraphs(label: string, xml: string): void {
	let root: XmlElement;
	try {
		root = parseXml(xml, { label }).documentElement;
	} catch {
		return;
	}
	const tree = parser.parse(xml) as XmlObject;
	partsCompared++;
	const visit = (element: XmlElement, node: unknown, trail: string) => {
		if (node === undefined) return;
		if (element.namespaceURI === NS.a && PARAGRAPH_ELEMENTS.has(element.localName)) {
			const objectNode = (typeof node === 'object' ? node : {}) as XmlObject;
			const dom = codecFamilies(parseDrawingParagraphProperties(element));
			const objectTree = objectParagraphProperties(objectNode);
			elementsCompared++;
			fieldsCompared += Object.keys(dom).length;
			if (canonical(objectTree) !== canonical(dom))
				readDifferences.push({
					label: `${label} ${trail}`,
					objectTree: canonical(objectTree),
					dom: canonical(dom),
				});
			const lineSpacing = dom.lineSpacing;
			const clamped =
				lineSpacing?.unit === 'percent' && (lineSpacing.value < 0.1 || lineSpacing.value > 5);
			if (!clamped) {
				savesCompared++;
				const saved = savedThroughPptx(objectNode);
				if (canonical(saved) !== canonical(savedFamilies(dom)))
					saveDifferences.push({
						label: `${label} ${trail}`,
						objectTree: canonical(saved),
						dom: canonical(savedFamilies(dom)),
					});
			}
			return;
		}
		const seen = new Map<string, number>();
		for (const child of elements(element)) {
			const index = seen.get(child.tagName) ?? 0;
			seen.set(child.tagName, index + 1);
			visit(child, objectChild(node, child.tagName, index), `${trail}/${child.tagName}#${index}`);
		}
	};
	visit(root, tree[root.tagName], root.tagName);
}

beforeAll(async () => {
	for (const deck of decks()) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(deck));
		} catch {
			continue; // The password-protected fixture is a CFB container, not a zip.
		}
		for (const name of Object.keys(zip.files).sort()) {
			if (zip.files[name]!.dir || !name.endsWith('.xml')) continue;
			const xml = await zip.file(name)!.async('string');
			if (!/pPr\b/u.test(xml)) continue;
			compareParagraphs(`${path.relative(CORE, deck)}!${name}`, xml);
		}
	}
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
