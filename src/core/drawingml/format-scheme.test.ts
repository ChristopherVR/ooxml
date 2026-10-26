import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';
import { parseThemeXml } from './theme';
import { drawingThemeXml } from './write-theme';
import { resolveDrawingColor } from './drawing-color';
import { themeDrawingColorTheme } from './theme-color';
import {
	drawingThemeFormatSchemeXml,
	parseThemeFormatScheme,
	resolveThemeEffectStyle,
	resolveThemeFillStyle,
	resolveThemeLineStyle,
	substituteDrawingPlaceholder,
} from './format-scheme';

const root = (xml: string) => parseXml(xml).documentElement;
const matrix = `<a:fmtScheme xmlns:a="${NS.a}" xmlns:v="urn:vendor" name="Authored" v:flag="keep"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill><a:grpFill vendor="opaque"/></a:fillStyleLst><a:lnStyleLst><a:ln w="012700" cmpd="dbl"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/><a:headEnd type="triangle"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle v:flag="keep"><a:effectLst><a:glow rad="12700"><a:schemeClr val="phClr"><a:alpha val="50000"/></a:schemeClr></a:glow></a:effectLst><a:scene3d/><a:sp3d/><v:extra/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:bgFillStyleLst><a:extLst><a:ext uri="keep"><v:matrix/></a:ext></a:extLst></a:fmtScheme>`;

// Compare namespace-aware structure while allowing redundant declarations inserted on imported roots.
function canonical(node: XmlElement): unknown {
	return [
		node.namespaceURI,
		node.localName,
		Array.from(node.attributes)
			.filter((attr) => attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
			.map((attr) => [attr.namespaceURI, attr.localName, attr.value])
			.sort(),
		elements(node).map(canonical),
	];
}

describe('shared theme format scheme', () => {
	it('retains all authored lists, opaque entries, 3D and extensions on round trip', () => {
		const scheme = parseThemeFormatScheme(root(matrix))!;
		expect(scheme.fillStyles[1]?.fill).toEqual({ kind: 'unsupported', element: 'grpFill' });
		expect(scheme.lineStyles[0]).toMatchObject({
			compound: 'dbl',
			join: { kind: 'round' },
			headEnd: { type: 'triangle' },
		});
		expect(canonical(root(drawingThemeFormatSchemeXml(scheme)))).toEqual(canonical(root(matrix)));
	});
	it('edits one list without resolving away any other authored theme colour', () => {
		const scheme = parseThemeFormatScheme(root(matrix))!;
		scheme.lineStyles[0]!.widthEmu = 25400;
		const saved = root(drawingThemeFormatSchemeXml(scheme));
		expect(first(first(saved, 'lnStyleLst', NS.a), 'ln', NS.a)?.getAttribute('w')).toBe('25400');
		for (const list of ['fillStyleLst', 'effectStyleLst', 'bgFillStyleLst', 'extLst'])
			expect(canonical(first(saved, list, NS.a)!)).toEqual(
				canonical(first(root(matrix), list, NS.a)!),
			);
	});
	it('resolves one-based references and phClr by composing transforms, keeping the source immutable', () => {
		const scheme = parseThemeFormatScheme(root(matrix))!;
		const before = structuredClone(scheme);
		const placeholder = {
			kind: 'scheme' as const,
			value: 'accent2',
			transforms: [{ name: 'alpha', value: '40000' }],
		};
		const fill = resolveThemeFillStyle(scheme, 1, placeholder)!;
		expect(fill).toMatchObject({
			kind: 'solid',
			color: {
				kind: 'scheme',
				value: 'accent2',
				transforms: [
					{ name: 'alpha', value: '40000' },
					{ name: 'shade', value: '50000' },
				],
			},
		});
		if (fill.kind !== 'solid') throw new Error('Missing fill');
		expect(resolveDrawingColor(fill.color, { scheme: () => '#FFFFFF' })?.alpha).toBe(0.4);
		expect(resolveThemeLineStyle(scheme, 1, placeholder)?.fill).toMatchObject({
			color: { value: 'accent2' },
		});
		expect(resolveThemeEffectStyle(scheme, 1, placeholder)?.effects[0]).toMatchObject({
			color: { value: 'accent2' },
		});
		expect(resolveThemeFillStyle(scheme, 1001)).toMatchObject({ color: { value: 'accent3' } });
		for (const index of [0, -1, 1.5, 1000, 1002, NaN, Infinity])
			expect(resolveThemeFillStyle(scheme, index)).toBeUndefined();
		expect(scheme).toEqual(before);
		expect(
			substituteDrawingPlaceholder({ kind: 'scheme', value: 'phClr', transforms: [] }),
		).toEqual({ kind: 'scheme', value: 'phClr', transforms: [] });
	});
	it('does not synthesize missing matrix lists in an override', () => {
		const xml = `<a:fmtScheme xmlns:a="${NS.a}" name="Partial"><a:lnStyleLst/></a:fmtScheme>`;
		expect(drawingThemeFormatSchemeXml(parseThemeFormatScheme(root(xml))!)).toBe(xml);
		expect(parseThemeXml(`<a:themeOverride xmlns:a="${NS.a}"/>`).formatScheme).toBeUndefined();
	});
});

describe('existing Office theme fixture parity', () => {
	it.each([
		['pptx/__tests__/fixtures/theme-styles.pptx', 'ppt/theme/theme1.xml'],
		['docx/__fixtures__/smartart.docx', 'word/theme/theme1.xml'],
		['xlsx/__fixtures__/excel-features.xlsx', 'xl/theme/theme1.xml'],
	])('reads and saves the native style matrix from %s', async (file, part) => {
		const zip = await JSZip.loadAsync(readFileSync(path.join(import.meta.dirname, '..', file)));
		const xml = await zip.file(part)!.async('string');
		const theme = parseThemeXml(xml);
		expect(theme.formatScheme?.lineStyles.length).toBeGreaterThan(0);
		const original = first(first(root(xml), 'themeElements', NS.a), 'fmtScheme', NS.a)!;
		expect(canonical(root(drawingThemeFormatSchemeXml(theme.formatScheme!)))).toEqual(
			canonical(original),
		);
		const saved = parseThemeXml(drawingThemeXml(theme));
		expect(canonical(root(drawingThemeFormatSchemeXml(saved.formatScheme!)))).toEqual(
			canonical(original),
		);
		const firstLine = resolveThemeLineStyle(saved.formatScheme!, 1, {
			kind: 'srgb',
			value: '123456',
			transforms: [],
		});
		if (firstLine?.fill?.kind === 'solid')
			expect(
				resolveDrawingColor(firstLine.fill.color, themeDrawingColorTheme(theme.colorScheme)),
			).toBeDefined();
		expect(buildXml(original)).toContain('phClr');
	});
});
