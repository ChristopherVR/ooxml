import { NS, buildXml, elements, first, parseXml } from '../../xml/index.js';
import type { ThemePalette } from '../model.js';
import { PALETTE_SLOTS, SCHEME_ORDER, parseTheme } from '../read/theme.js';
import { XML_HEADER, escapeAttr } from './xml-out.js';

const colorOf = (theme: ThemePalette, slot: string): string =>
	(theme.colors[PALETTE_SLOTS.indexOf(slot as (typeof PALETTE_SLOTS)[number])] ?? '000000')
		.slice(-6)
		.toUpperCase();

/** A complete Office theme part built from the model palette and fonts. */
export function defaultThemeXml(theme: ThemePalette): string {
	const scheme = SCHEME_ORDER.map(
		(slot) => `<a:${slot}><a:srgbClr val="${colorOf(theme, slot)}"/></a:${slot}>`,
	).join('');
	const font = (kind: string, face: string) =>
		`<a:${kind}><a:latin typeface="${escapeAttr(face)}"/><a:ea typeface=""/><a:cs typeface=""/></a:${kind}>`;
	const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
	const line = (w: number) =>
		`<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
	const effect = '<a:effectStyle><a:effectLst/></a:effectStyle>';
	return (
		`${XML_HEADER}<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements>` +
		`<a:clrScheme name="Office">${scheme}</a:clrScheme>` +
		`<a:fontScheme name="Office">${font('majorFont', theme.majorFont)}${font('minorFont', theme.minorFont)}</a:fontScheme>` +
		`<a:fmtScheme name="Office"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst>` +
		`<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
		`<a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst>` +
		`<a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme>` +
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
	for (const slot of elements(scheme)) {
		for (const child of elements(slot)) slot.removeChild(child);
		const color = doc.createElementNS(NS.a, 'a:srgbClr');
		color.setAttribute('val', colorOf(theme, slot.localName ?? ''));
		slot.appendChild(color);
	}
	const fonts = first(themeElements, 'fontScheme', NS.a);
	for (const [kind, face] of [
		['majorFont', theme.majorFont],
		['minorFont', theme.minorFont],
	] as const)
		first(first(fonts, kind, NS.a), 'latin', NS.a)?.setAttribute('typeface', face);
	return `${XML_HEADER}${buildXml(doc).replace(/^<\?xml[^>]*\?>\s*/, '')}`;
}
