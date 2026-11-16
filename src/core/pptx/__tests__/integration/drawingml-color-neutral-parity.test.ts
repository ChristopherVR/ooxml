import JSZip from 'jszip';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import { declaredDrawingColorAlpha } from '../../../drawingml/color-transforms';
import { parseDrawingColor, resolveDrawingColor } from '../../../drawingml/drawing-color';
import { parseTheme } from '../../../drawingml/theme';
import { NS, elements, parseXml, type XmlElement } from '../../../xml/index';
import {
	resolveObjectColorChoice,
	resolveObjectColorOpacity,
} from '../../core/color/color-transforms';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { themeSchemesFromDrawingTheme } from '../../core/core/runtime/theme-part-model';
import { THEME_COLOR_SCHEME_KEYS } from '../../core/core/runtime/theme-scheme-edit';
import { DEFAULT_SCHEME_COLOR_MAP } from '../../core/constants';
import type { XmlObject } from '../../core/types';

/**
 * DrawingML colour and theme parity (docs/agnostic-core-plan.md step 3): the pptx colour codec
 * and theme part reader resolve through the shared `drawingml` reader and resolver, and this
 * harness proves the object-tree path and the DOM path agree over every part of every committed
 * deck (`pptx/__tests__/fixtures`, including the `e2e` snapshot).
 *
 * - Colours: every colour element (`a:srgbClr`, `a:schemeClr`, ...) is read twice, through the
 *   pptx object tree (`resolveObjectColorChoice`, what `parseColorChoice` runs) and through the
 *   DOM (`parseDrawingColor` + `resolveDrawingColor`), paired by element path (same-name sibling
 *   index), and the hex values must be equal.
 * - Themes: every `a:theme` part's colour slots and fonts, as the runtime keeps them
 *   (`themeSchemesFromDrawingTheme` over `parseTheme`), must equal an independent object-tree
 *   reading of the same part (slot colours through the codec, typefaces trimmed, east Asian and
 *   complex script faces falling back to latin).
 *
 * - Opacity: the same pairs must declare the same alpha (`resolveObjectColorOpacity`, what the
 *   codec's `extractColorOpacity` and `parseDrawingColorOpacity` run, against the resolver's alpha
 *   for the DOM colour, `undefined` when no alpha transform is written).
 *
 * Before the opacity switch, the previous readers (the codec's `extractColorOpacity`, the core and
 * UI `parseDrawingColorOpacity` and the UI's own colour-choice reader) were compared with the new
 * ones over 180,499 colour holders (15,102 declaring alpha): identical. They differed only where the
 * corpus has no case: an `a:alphaMod` above 100% (the old readers clamped the modulator first), a
 * repeated alpha transform (ignored, now the first applies) and, in the UI reader, a negative
 * `a:alphaOff` (clamped to 0) and a `sysClr` without `lastClr` (now the system table).
 *
 * Before the switch, the previous object-tree codec and theme reader were compared with the new
 * ones over the same corpus (180,499 colour holders in 15,928 parts, 383 theme parts): equal
 * except for 42 colours in `three-d-smartart.pptx` drawings that repeat `a:tint`, which the old
 * codec skipped and the shared resolver applies (first occurrence), as the DOM reader does.
 */

const CORE = path.resolve(__dirname, '../../..');
const COLOR_ELEMENTS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);
const SCHEME: Record<string, string> = {
	dk1: '#101010',
	lt1: '#F0F0F0',
	dk2: '#203040',
	lt2: '#E0E8F0',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	hlink: '#0563C1',
	folhlink: '#954F72',
	tx1: '#111111',
	bg1: '#FEFEFE',
	tx2: '#44546A',
	bg2: '#E7E6E6',
	phclr: '#4472C4',
};
const scheme = (key: string) => SCHEME[key];
const domTheme = { scheme: (name: string) => scheme(name.toLowerCase()) };
const parser = new PptxRuntimeDependencyFactory().createParser();

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

interface Difference {
	label: string;
	objectTree: string | undefined;
	dom: string | undefined;
}

const differences: Difference[] = [];
let colorsCompared = 0;
let alphasDeclared = 0;
let partsCompared = 0;
let partsRejected = 0;
let themesCompared = 0;
/**
 * DrawingML elements under a DOM element the object tree does not hold at the same place (the
 * `a14:m` maths runs of one e2e deck); they are counted, not compared.
 */
let colorsUnpaired = 0;

function compareColors(label: string, xml: string): void {
	let root: XmlElement;
	try {
		root = parseXml(xml, { label }).documentElement;
	} catch {
		partsRejected++;
		return;
	}
	const tree = parser.parse(xml) as XmlObject;
	partsCompared++;
	const visit = (element: XmlElement, node: unknown, trail: string) => {
		if (node === undefined) {
			colorsUnpaired += element.getElementsByTagNameNS(NS.a, '*').length;
			return;
		}
		if (element.namespaceURI === NS.a && COLOR_ELEMENTS.has(element.localName)) {
			const holder = { [element.tagName]: typeof node === 'object' ? node : {} } as XmlObject;
			const objectTree = resolveObjectColorChoice(holder, scheme);
			const color = parseDrawingColor(element);
			const dom = color ? resolveDrawingColor(color, domTheme)?.hex : undefined;
			colorsCompared++;
			if (objectTree !== dom) differences.push({ label: `${label} ${trail}`, objectTree, dom });
			const objectAlpha = resolveObjectColorOpacity(holder);
			const domAlpha = color ? declaredDrawingColorAlpha(color.transforms) : undefined;
			if (objectAlpha !== undefined) alphasDeclared++;
			if (objectAlpha !== domAlpha)
				differences.push({
					label: `${label} ${trail} alpha`,
					objectTree: String(objectAlpha),
					dom: String(domAlpha),
				});
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

/** An object-tree reading of a theme part's slots and fonts, independent of the shared reader. */
function objectTreeTheme(xml: string) {
	const root = (parser.parse(xml) as XmlObject)['a:theme'] as XmlObject | undefined;
	const holder = root?.['a:themeElements'] as XmlObject | undefined;
	const colors = holder?.['a:clrScheme'] as XmlObject | undefined;
	const fonts = holder?.['a:fontScheme'] as XmlObject | undefined;
	const colorMap: Record<string, string> = { ...DEFAULT_SCHEME_COLOR_MAP };
	for (const key of THEME_COLOR_SCHEME_KEYS) {
		const hex = resolveObjectColorChoice(
			colors?.[`a:${key}`] as XmlObject | undefined,
			() => undefined,
		);
		if (hex) colorMap[key] = hex;
	}
	const fontMap: Record<string, string> = {};
	const face = (node: unknown) =>
		String((node as XmlObject | undefined)?.['@_typeface'] ?? '').trim() || undefined;
	for (const [prefix, tag] of [
		['mj', 'a:majorFont'],
		['mn', 'a:minorFont'],
	] as const) {
		const group = fonts?.[tag] as XmlObject | undefined;
		const latin = face(group?.['a:latin']);
		if (latin) for (const slot of ['lt', 'ea', 'cs']) fontMap[`${prefix}-${slot}`] = latin;
		const eastAsia = face(group?.['a:ea']);
		if (eastAsia) fontMap[`${prefix}-ea`] = eastAsia;
		const complexScript = face(group?.['a:cs']);
		if (complexScript) fontMap[`${prefix}-cs`] = complexScript;
	}
	return { name: String(root?.['@_name'] ?? '').trim() || undefined, colorMap, fontMap };
}

function compareTheme(label: string, xml: string): void {
	const expected = objectTreeTheme(xml);
	const schemes = themeSchemesFromDrawingTheme(
		parseTheme(parseXml(xml, { label })),
		DEFAULT_SCHEME_COLOR_MAP,
	);
	themesCompared++;
	const actual = { name: schemes.name, colorMap: schemes.colorMap, fontMap: schemes.fontMap };
	if (JSON.stringify(actual) !== JSON.stringify(expected))
		differences.push({ label, objectTree: JSON.stringify(expected), dom: JSON.stringify(actual) });
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
			if (!/Clr\b|clrScheme|fontScheme/u.test(xml)) continue;
			const label = `${path.relative(CORE, deck)}!${name}`;
			compareColors(label, xml);
			if (/^ppt\/theme\/theme\d*\.xml$/u.test(name)) compareTheme(label, xml);
		}
	}
}, 600_000);

describe('DrawingML colours and themes: pptx object tree vs shared DOM reader', () => {
	it('compares a real corpus', () => {
		expect(partsCompared).toBeGreaterThan(3000);
		expect(colorsCompared).toBeGreaterThan(150_000);
		expect(alphasDeclared).toBeGreaterThan(10_000);
		expect(themesCompared).toBeGreaterThan(300);
		expect(partsRejected).toBeLessThan(10);
		expect(colorsUnpaired).toBeLessThan(500);
	});

	it('agrees on every colour element and every theme part', () => {
		expect(differences.slice(0, 20)).toEqual([]);
	});
});
