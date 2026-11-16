import JSZip from 'jszip';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import { NS, first, parseXml } from '../../../xml/index';
import { masterColorMapOf } from '../../core/core/runtime/PptxHandlerRuntimeThemeLoading';
import type { ThemePartModel } from '../../core/core/runtime/theme-part-model';
import type { PptxThemeFormatScheme, ShapeStyle, XmlObject } from '../../core/types';
import { LegacyThemeFormatSchemeRuntime } from '../helpers/legacy-theme-format-scheme';

/**
 * DrawingML theme format scheme parity (docs/agnostic-core-plan.md step 3). The pptx format scheme
 * reader, the `fillRef`/`lnRef`/`effectRef` resolution with its `phClr` substitution and the
 * master `p:clrMap` reader run on the shared DrawingML model (`parseThemeFormatScheme`,
 * `resolveTheme*Style`, `parseThemeColorMap`). This harness compares them with the object-tree
 * code they replaced (kept as a frozen reference in `helpers/legacy-theme-format-scheme.ts`) over
 * every committed deck (`pptx/__tests__/fixtures`, including the `e2e` snapshot):
 *
 * - every `a:fmtScheme` of every theme and theme override part, as the pptx view;
 * - every style reference of every slide, layout and master, resolved against the deck's first
 *   theme;
 * - every master's `p:clrMap`;
 * - every theme part written back from the runtime's state (`buildThemeXml`, a dirty theme).
 *
 * The neutral halves the new view adds (`fill`, `line`, `effects`) are not compared. The one
 * intended difference is listed in `EXPECTED_SCHEME_DIFFERENCES`.
 */

const CORE = path.resolve(__dirname, '../../..');

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

class ParityRuntime extends LegacyThemeFormatSchemeRuntime {
	public readThemePart(xml: string): ThemePartModel {
		return this.parseThemePartXml(xml);
	}
	public parse(xml: string): XmlObject {
		return this.parser.parse(xml) as XmlObject;
	}
	public newFormatScheme(node: XmlObject, xml: string): PptxThemeFormatScheme {
		return this.parseFormatScheme(node, first(parseXml(xml).documentElement, 'fmtScheme', NS.a));
	}
	public useTheme(colorMap: Record<string, string>, scheme: PptxThemeFormatScheme): void {
		this.themeColorMap = { ...colorMap };
		this.themeFormatScheme = scheme;
	}
	public resolveNew(kind: RefKind, node: XmlObject): ShapeStyle {
		const style: ShapeStyle = {};
		if (kind === 'a:fillRef') this.resolveThemeFillRef(node, style);
		else if (kind === 'a:lnRef') this.resolveThemeLineRef(node, style);
		else this.resolveThemeEffectRef(node, style);
		return style;
	}
	/** One theme part written back by the old and the new writer, as object trees. */
	public writeTheme(themePath: string, xml: string): { legacy: unknown; next: unknown } {
		const model = this.parseThemePartXml(xml);
		this.cacheThemePartModel(themePath, xml, model);
		this.themeColorMap = { ...model.colorMap };
		this.themeFontMap = { ...model.fontMap };
		this.globalThemeColorMapSnapshot = { ...model.colorMap };
		this.globalThemeFontMapSnapshot = { ...model.fontMap };
		const tree = (text: string | undefined) => (text === undefined ? undefined : this.parse(text));
		return {
			legacy: tree(this.legacyBuildThemeXml(themePath, undefined)),
			next: tree(this.buildThemeXml(themePath, undefined)),
		};
	}
	public resolveLegacy(kind: RefKind, node: XmlObject): ShapeStyle {
		const style: ShapeStyle = {};
		if (kind === 'a:fillRef') this.legacyResolveThemeFillRef(node, style);
		else if (kind === 'a:lnRef') this.legacyResolveThemeLineRef(node, style);
		else this.legacyResolveThemeEffectRef(node, style);
		return style;
	}
}

type RefKind = 'a:fillRef' | 'a:lnRef' | 'a:effectRef';
const REF_KINDS: readonly RefKind[] = ['a:fillRef', 'a:lnRef', 'a:effectRef'];

/** The view without the neutral halves the new reader adds. */
function legacyShape(scheme: PptxThemeFormatScheme | undefined): unknown {
	if (!scheme) return undefined;
	const strip = (entries: object[], key: string) =>
		entries.map((entry) => {
			const copy = { ...entry } as Record<string, unknown>;
			delete copy[key];
			return copy;
		});
	return {
		name: scheme.name,
		fillStyles: strip(scheme.fillStyles, 'fill'),
		lineStyles: strip(scheme.lineStyles, 'line'),
		effectStyles: strip(scheme.effectStyles, 'effects'),
		backgroundFillStyles: strip(scheme.backgroundFillStyles, 'fill'),
	};
}

/** Every style reference node under an object tree. */
function collectRefs(node: unknown, out: Array<[RefKind, XmlObject]>): void {
	if (Array.isArray(node)) {
		for (const item of node) collectRefs(item, out);
		return;
	}
	if (!node || typeof node !== 'object') return;
	for (const [key, value] of Object.entries(node as XmlObject)) {
		if ((REF_KINDS as readonly string[]).includes(key)) {
			for (const ref of Array.isArray(value) ? value : [value]) {
				if (ref && typeof ref === 'object') out.push([key as RefKind, ref as XmlObject]);
			}
		} else if (typeof value === 'object') collectRefs(value, out);
	}
}

/** The object-tree `p:clrMap` reader the new one replaced (values lower-cased, unvalidated). */
function legacyMasterColorMap(tree: XmlObject): Record<string, string> | undefined {
	const clrMap = (tree['p:sldMaster'] as XmlObject | undefined)?.['p:clrMap'] as
		| XmlObject
		| undefined;
	if (!clrMap) return undefined;
	const map: Record<string, string> = {};
	for (const key of [
		'bg1',
		'tx1',
		'bg2',
		'tx2',
		'accent1',
		'accent2',
		'accent3',
		'accent4',
		'accent5',
		'accent6',
		'hlink',
		'folHlink',
	]) {
		const value = String(clrMap[`@_${key}`] || '')
			.trim()
			.toLowerCase();
		if (value) map[key] = value;
	}
	return Object.keys(map).length > 0 ? map : undefined;
}

/** The paths at which two values differ (`fillStyles.1.color`), for the failure report. */
function differingPaths(a: unknown, b: unknown, at = ''): string[] {
	if (Object.is(a, b)) return [];
	const isObject = (value: unknown) => value !== null && typeof value === 'object';
	if (!isObject(a) || !isObject(b) || Array.isArray(a) !== Array.isArray(b)) {
		return [`${at}: ${JSON.stringify(a)?.slice(0, 80)} -> ${JSON.stringify(b)?.slice(0, 80)}`];
	}
	const keys = new Set([...Object.keys(a as object), ...Object.keys(b as object)]);
	return [...keys].flatMap((key) =>
		differingPaths(
			(a as Record<string, unknown>)[key],
			(b as Record<string, unknown>)[key],
			at ? `${at}.${key}` : key,
		),
	);
}

/** Deep equality as `toEqual` reads it (an undefined property equals a missing one). */
const same = (a: unknown, b: unknown): boolean => differingPaths(a, b).length === 0;

interface Difference {
	label: string;
	legacy: unknown;
	next: unknown;
}

const schemeDifferences: Difference[] = [];
const refDifferences: Difference[] = [];
const mapDifferences: Difference[] = [];
const writeDifferences: Difference[] = [];
let themesWritten = 0;
let schemesCompared = 0;
let refsCompared = 0;
let mapsCompared = 0;
let decksRead = 0;

/**
 * Intended differences. One e2e deck's theme lists an `a:blipFill` in its fill and background
 * fill lists: the old reader dropped it (so the entries after it moved up one index), the shared
 * reader keeps it at its index as a `picture` entry.
 */
const EXPECTED_SCHEME_DIFFERENCES = [
	'pptx/__tests__/fixtures/e2e/absolute-path-rels.pptx ppt/theme/theme11.xml',
];

beforeAll(async () => {
	const runtime = new ParityRuntime();
	for (const deck of decks()) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(deck));
		} catch {
			continue;
		}
		decksRead++;
		const name = path.relative(CORE, deck).replace(/\\/g, '/');
		const parts = Object.keys(zip.files).sort();
		let deckTheme:
			| {
					colorMap: Record<string, string>;
					legacy: PptxThemeFormatScheme;
					next: PptxThemeFormatScheme;
			  }
			| undefined;
		for (const part of parts) {
			const isTheme = /^ppt\/theme\/theme\d+\.xml$/i.test(part);
			const isOverride = /^ppt\/theme\/themeOverride\d+\.xml$/i.test(part);
			if (!isTheme && !isOverride) continue;
			const xml = await zip.file(part)!.async('string');
			const tree = runtime.parse(xml);
			const root = (tree['a:theme'] ?? tree['a:themeOverride']) as XmlObject | undefined;
			const fmt = (
				isTheme
					? (root?.['a:themeElements'] as XmlObject | undefined)?.['a:fmtScheme']
					: root?.['a:fmtScheme']
			) as XmlObject | undefined;
			if (!fmt) continue;
			let next: PptxThemeFormatScheme | undefined;
			let colorMap: Record<string, string> = {};
			if (isTheme) {
				// The scheme's colours resolve through the theme in force: read the part's own first.
				colorMap = runtime.readThemePart(xml).colorMap;
				runtime.useTheme(colorMap, undefined as never);
				next = runtime.readThemePart(xml).formatScheme;
			} else {
				runtime.useTheme(deckTheme?.colorMap ?? {}, undefined as never);
				next = runtime.newFormatScheme(fmt, xml);
			}
			const legacy = runtime.legacyParseFormatScheme(fmt);
			schemesCompared++;
			if (!same(legacyShape(next), legacyShape(legacy))) {
				schemeDifferences.push({
					label: `${name} ${part}`,
					legacy: legacyShape(legacy),
					next: legacyShape(next),
				});
			}
			if (isTheme && !deckTheme && next) deckTheme = { colorMap, legacy, next };
			if (isTheme) {
				const written = new ParityRuntime().writeTheme(part, xml);
				themesWritten++;
				if (!same(written.next, written.legacy)) {
					writeDifferences.push({ label: `${name} ${part}`, ...written });
				}
			}
		}
		for (const part of parts) {
			if (
				!/^ppt\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|diagrams)\/[^/]+\.xml$/i.test(
					part,
				)
			)
				continue;
			const xml = await zip.file(part)!.async('string');
			if (/slideMasters/.test(part)) {
				const legacy = legacyMasterColorMap(runtime.parse(xml));
				const next = masterColorMapOf(xml);
				mapsCompared++;
				const lowered =
					next && Object.fromEntries(Object.entries(next).map(([k, v]) => [k, v.toLowerCase()]));
				if (!same(lowered, legacy)) {
					mapDifferences.push({ label: `${name} ${part}`, legacy, next });
				}
			}
			if (!deckTheme) continue;
			const refs: Array<[RefKind, XmlObject]> = [];
			collectRefs(runtime.parse(xml), refs);
			for (const [kind, node] of refs) {
				runtime.useTheme(deckTheme.colorMap, deckTheme.legacy);
				const legacy = runtime.resolveLegacy(kind, node);
				runtime.useTheme(deckTheme.colorMap, deckTheme.next);
				const next = runtime.resolveNew(kind, node);
				refsCompared++;
				if (!same(next, legacy)) {
					refDifferences.push({ label: `${name} ${part} ${kind}`, legacy, next });
				}
			}
		}
	}
}, 600_000);

describe('DrawingML theme format scheme parity over the committed decks', () => {
	it('reads the committed corpus', () => {
		expect(decksRead).toBeGreaterThan(300);
		expect(schemesCompared).toBeGreaterThan(350);
		expect(refsCompared).toBeGreaterThan(60_000);
		expect(mapsCompared).toBeGreaterThan(350);
		expect(themesWritten).toBeGreaterThan(350);
	});

	it('writes every theme part back as the object-tree writer did', () => {
		for (const d of writeDifferences.slice(0, 8))
			console.log(d.label, differingPaths(d.legacy, d.next).slice(0, 6));
		expect(writeDifferences.map((d) => d.label)).toEqual([]);
	});

	it('reads every format scheme as the object-tree reader did', () => {
		for (const d of schemeDifferences.slice(0, 8))
			console.log(d.label, differingPaths(d.legacy, d.next).slice(0, 6));
		expect(schemeDifferences.map((d) => d.label)).toEqual(EXPECTED_SCHEME_DIFFERENCES);
		const [blip] = schemeDifferences;
		expect(differingPaths(blip?.legacy, blip?.next).map((at) => at.split(':')[0])).toEqual([
			'fillStyles.2',
			'backgroundFillStyles.2',
		]);
	});

	it('resolves every style reference as the object-tree resolution did', () => {
		for (const d of refDifferences.slice(0, 8))
			console.log(d.label, differingPaths(d.legacy, d.next).slice(0, 6));
		expect(refDifferences.map((d) => d.label)).toEqual([]);
	});

	it('reads every master colour map as the object-tree reader did', () => {
		for (const d of mapDifferences.slice(0, 5)) console.log(JSON.stringify(d));
		expect(mapDifferences).toEqual([]);
	});
});
