import type JSZip from 'jszip';
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';
import { drawingThemeFormatSchemeXml, type ThemeFormatScheme } from '../drawingml/format-scheme';
import { parseDrawingColorIn } from '../drawingml/drawing-color';
import { patchDrawingColorIn } from '../drawingml/patch-color';
import { THEME_COLOR_SLOTS } from '../drawingml/theme-model';
import { drawingThemeXml } from '../drawingml/write-theme';
import { themeMatrix } from '../drawingml/theme-matrix';
import type { DrawingColor } from '../drawingml/types';
import { documentPartPath, parseRelationships } from './package-parts';
import { applyColorSchemeMapping } from './settings';
import type { ThemeCatalog } from './theme-model';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts';
import { writeWordThemeMedia } from './theme-media';

const THEME_PATH = 'word/theme/theme1.xml';
const FONT_ROLES = ['major', 'minor'] as const;
const FONT_FIELDS = [
	['latin', 'latin'],
	['eastAsia', 'ea'],
	['complexScript', 'cs'],
] as const;
const THEME_FALLBACKS = [
	'000000',
	'FFFFFF',
	'44546A',
	'E7E6E6',
	'4472C4',
	'ED7D31',
	'A5A5A5',
	'FFC000',
	'5B9BD5',
	'70AD47',
	'0563C1',
	'954F72',
];

function rgb(hex: string): DrawingColor {
	if (!/^#?[0-9a-f]{6}$/i.test(hex)) throw new Error('Theme colors must be six-digit RGB values.');
	return { kind: 'srgb', value: hex.replace(/^#/, ''), transforms: [] };
}

export function wordThemeXml(theme: ThemeCatalog): string {
	return drawingThemeXml({
		name: theme.name ?? 'Office Theme',
		colorScheme: {
			...(theme.colorSchemeName !== undefined && { name: theme.colorSchemeName }),
			colors: Object.fromEntries(
				THEME_COLOR_SLOTS.map((slot, index) => [
					slot,
					rgb(theme.colors[slot] ?? THEME_FALLBACKS[index] ?? '000000'),
				]),
			),
		},
		fontScheme: {
			...(theme.fontSchemeName !== undefined && { name: theme.fontSchemeName }),
			major: { scripts: {}, ...theme.fonts.major },
			minor: { scripts: {}, ...theme.fonts.minor },
		},
		formatScheme: theme.formatScheme ?? themeMatrix(),
	});
}

function ensure(parent: XmlElement, local: string, order: string[]): XmlElement {
	const old = first(parent, local, NS.a);
	if (old) return old;
	const next = parent.ownerDocument.createElementNS(NS.a, `a:${local}`);
	const later = order.slice(order.indexOf(local) + 1);
	parent.insertBefore(
		next,
		elements(parent).find(
			(node) => node.namespaceURI === NS.a && later.includes(node.localName ?? ''),
		) ?? null,
	);
	return next;
}

function formatStyles(scheme: ThemeFormatScheme) {
	const { name: _name, sourceXml: _source, ...styles } = scheme;
	return styles;
}

function patchFormatScheme(
	node: XmlElement,
	next: ThemeFormatScheme | undefined,
	base: ThemeFormatScheme | undefined,
) {
	if (next && base && JSON.stringify(formatStyles(next)) === JSON.stringify(formatStyles(base))) {
		if (next.name === undefined) node.removeAttribute('name');
		else node.setAttribute('name', next.name);
		return;
	}
	const sourceXml = next?.sourceXml ?? base?.sourceXml;
	const xml = next
		? drawingThemeFormatSchemeXml({ ...next, ...(sourceXml && { sourceXml }) })
		: drawingThemeFormatSchemeXml(themeMatrix());
	const format = parseXml(
		xml.includes('xmlns:a=') ? xml : xml.replace('<a:fmtScheme', `<a:fmtScheme xmlns:a="${NS.a}"`),
	).documentElement;
	node.parentNode?.replaceChild(node.ownerDocument.importNode(format, true), node);
}

function patchTheme(xml: string, next: ThemeCatalog, base: ThemeCatalog): string {
	const doc = parseXml(xml);
	const theme = ensure(doc.documentElement, 'themeElements', [
		'themeElements',
		'objectDefaults',
		'extraClrSchemeLst',
		'custClrLst',
		'extLst',
	]);
	const order = ['clrScheme', 'fontScheme', 'fmtScheme', 'extLst'];
	for (const [key, local] of [
		['name', ''],
		['colorSchemeName', 'clrScheme'],
		['fontSchemeName', 'fontScheme'],
	] as const) {
		if (next[key] === base[key]) continue;
		const node = local ? ensure(theme, local, order) : doc.documentElement;
		node.setAttribute('name', next[key] ?? 'Custom');
	}
	for (const [index, slot] of THEME_COLOR_SLOTS.entries()) {
		const value = next.colors[slot];
		if (value === base.colors[slot]) continue;
		const color = rgb(value ?? THEME_FALLBACKS[index] ?? '000000');
		const scheme = ensure(theme, 'clrScheme', order);
		const entry = ensure(scheme, slot, [...THEME_COLOR_SLOTS, 'extLst']);
		const source = parseDrawingColorIn(entry);
		// An explicit RGB edit must be used even on machines whose system color differs.
		patchDrawingColorIn(entry, { ...color, transforms: source?.transforms ?? [] });
	}
	for (const role of FONT_ROLES) {
		if (JSON.stringify(next.fonts[role].scripts) !== JSON.stringify(base.fonts[role].scripts)) {
			const scheme = ensure(theme, 'fontScheme', order);
			const fonts = ensure(scheme, `${role}Font`, ['majorFont', 'minorFont', 'extLst']);
			const ext = first(fonts, 'extLst', NS.a) ?? null;
			for (const font of elements(fonts))
				if (font.localName === 'font' && font.namespaceURI === NS.a) fonts.removeChild(font);
			for (const [script, face] of Object.entries(next.fonts[role].scripts ?? {})) {
				const node = doc.createElementNS(NS.a, 'a:font');
				node.setAttribute('script', script);
				node.setAttribute('typeface', face);
				fonts.insertBefore(node, ext);
			}
		}
		for (const [key, local] of FONT_FIELDS) {
			if (next.fonts[role][key] === base.fonts[role][key]) continue;
			const scheme = ensure(theme, 'fontScheme', order);
			const fonts = ensure(scheme, `${role}Font`, ['majorFont', 'minorFont', 'extLst']);
			const font = ensure(fonts, local, ['latin', 'ea', 'cs', 'font', 'extLst']);
			font.setAttribute('typeface', next.fonts[role][key] ?? '');
		}
	}
	if (JSON.stringify(next.formatScheme) !== JSON.stringify(base.formatScheme))
		patchFormatScheme(ensure(theme, 'fmtScheme', order), next.formatScheme, base.formatScheme);
	return buildXml(doc);
}

export async function applyWordTheme(
	zip: JSZip,
	next: ThemeCatalog | undefined,
	base: ThemeCatalog | undefined,
): Promise<void> {
	if (JSON.stringify(next) === JSON.stringify(base)) return;
	if (!next) throw new Error('Removing a Word theme is unsupported.');
	const rels = parseRelationships(await zip.file('word/_rels/document.xml.rels')?.async('string'));
	const relatedPath = documentPartPath(rels, 'theme', '');
	const path = relatedPath || THEME_PATH;
	const file = zip.file(path);
	if (!file || !base) {
		zip.file(path, wordThemeXml(next));
		await ensureContentTypeOverride(
			zip,
			path,
			'application/vnd.openxmlformats-officedocument.theme+xml',
		);
		if (!relatedPath)
			await ensureDocumentRelationship(
				zip,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
				'theme/theme1.xml',
			);
	} else if (
		JSON.stringify({ ...next, colorMapping: {} }) !== JSON.stringify({ ...base, colorMapping: {} })
	) {
		zip.file(path, patchTheme(await file.async('string'), next, base));
	}
	await writeWordThemeMedia(zip, path, next);
	await applyColorSchemeMapping(
		zip,
		next.colorMapping,
		base?.colorMapping,
		documentPartPath(rels, 'settings', ''),
	);
}
