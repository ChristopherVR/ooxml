import { NS, buildXml, elements, first, parseXml } from '../../xml/index';
import type { ThemePalette } from '../model';
import { drawingThemeFormatSchemeXml } from '../../drawingml/format-scheme';
import { OFFICE_FORMAT_SCHEME_XML } from '../office-format-scheme';
import { THEME_COLOR_SLOTS } from '../../drawingml/theme-model';
import { PALETTE_SLOTS, parseTheme } from '../read/theme';
import { XML_HEADER, escapeAttr } from './xml-out';

const colorOf = (theme: ThemePalette, slot: string): string =>
	(theme.colors[PALETTE_SLOTS.indexOf(slot as (typeof PALETTE_SLOTS)[number])] ?? '000000')
		.slice(-6)
		.toUpperCase();

/** A complete Office theme part built from the model palette and fonts. */
export function defaultThemeXml(theme: ThemePalette): string {
	const scheme = THEME_COLOR_SLOTS.map(
		(slot) => `<a:${slot}><a:srgbClr val="${colorOf(theme, slot)}"/></a:${slot}>`,
	).join('');
	const font = (kind: string, face: string) =>
		`<a:${kind}><a:latin typeface="${escapeAttr(face)}"/><a:ea typeface=""/><a:cs typeface=""/></a:${kind}>`;
	return (
		`${XML_HEADER}<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements>` +
		`<a:clrScheme name="Office">${scheme}</a:clrScheme>` +
		`<a:fontScheme name="Office">${font('majorFont', theme.majorFont)}${font('minorFont', theme.minorFont)}</a:fontScheme>` +
		(theme.formatScheme
			? drawingThemeFormatSchemeXml(theme.formatScheme)
			: OFFICE_FORMAT_SCHEME_XML) +
		'</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
	);
}

/**
 * The theme part to write: the source bytes when the palette and fonts are unchanged, the
 * source with its colour scheme and Latin fonts patched otherwise, or a new default theme.
 */
export function themePart(theme: ThemePalette, sourceXml: string | undefined): string | undefined {
	if (!sourceXml) return defaultThemeXml(theme);
	const current = parseTheme(sourceXml);
	if (
		current.majorFont === theme.majorFont &&
		current.minorFont === theme.minorFont &&
		current.colors.join() === theme.colors.map((c) => c.slice(-6).toUpperCase()).join()
	)
		return undefined;
	const doc = parseXml(sourceXml, { label: 'XLSX theme' });
	const themeElements = first(doc.documentElement, 'themeElements', NS.a);
	const scheme = first(themeElements, 'clrScheme', NS.a);
	if (!scheme) return defaultThemeXml(theme);
	const slotSet = new Set<string>(THEME_COLOR_SLOTS);
	for (const slot of elements(scheme)) {
		if (!slotSet.has(slot.localName ?? '')) continue;
		for (const child of elements(slot)) slot.removeChild(child);
		const color = doc.createElementNS(NS.a, 'a:srgbClr');
		color.setAttribute('val', colorOf(theme, slot.localName ?? ''));
		slot.appendChild(color);
	}
	const fonts = first(themeElements, 'fontScheme', NS.a);
	for (const [kind, face] of [
		['majorFont', theme.majorFont],
		['minorFont', theme.minorFont],
	] as const) {
		const group = first(fonts, kind, NS.a);
		if (!group) continue;
		let latin = first(group, 'latin', NS.a);
		if (!latin) {
			latin = doc.createElementNS(NS.a, 'a:latin');
			group.insertBefore(latin, group.firstChild);
		}
		latin.setAttribute('typeface', face);
	}
	return `${XML_HEADER}${buildXml(doc).replace(/^<\?xml[^>]*\?>\s*/, '')}`;
}
