import { children, first, getW, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';
import { isThemeColorToken, resolveThemeColorReference, themeByteToFraction } from './theme-color';
import type { ThemeCatalog, ThemeColorReference } from './theme-model';

/** `w:background` literal colours are six hex digits. */
const HEX = /^[0-9a-fA-F]{6}$/;

/** The page colour of `w:document/w:background`, as uppercase hex without `#`. */
export function parsePageBackground(
	root: XmlElement | null | undefined,
	theme?: ThemeCatalog,
): string | undefined {
	const background = first(root, 'background');
	const token = getW(background, 'themeColor');
	if (theme && token && isThemeColorToken(token)) {
		const reference: ThemeColorReference = { token };
		const tint = themeByteToFraction(getW(background, 'themeTint'));
		const shade = themeByteToFraction(getW(background, 'themeShade'));
		if (tint !== undefined) reference.tint = tint;
		if (shade !== undefined) reference.shade = shade;
		const resolved = resolveThemeColorReference(reference, theme);
		if (resolved) return resolved.replace(/^#/, '').toUpperCase();
	}
	const color = getW(background, 'color');
	return color && HEX.test(color) ? color.toUpperCase() : undefined;
}

/**
 * Sets or clears an explicitly changed page colour. The caller skips unchanged model colours.
 * A literal replaces theme attributes even when its RGB matches the source's cached value.
 */
export function applyPageBackground(doc: XmlDocument, color: string | undefined): void {
	const root = doc.documentElement as XmlElement;
	const existing = first(root, 'background');
	const hex = color?.replace(/^#/, '').toUpperCase() ?? '';
	if (!HEX.test(hex)) {
		if (existing) root.removeChild(existing);
		return;
	}
	if (existing && !getW(existing, 'themeColor') && getW(existing, 'color')?.toUpperCase() === hex)
		return;
	const element = makeW(doc, 'background');
	element.setAttributeNS(WORD_NS, 'w:color', hex);
	if (existing) root.replaceChild(element, existing);
	else root.insertBefore(element, children(root, 'body')[0] ?? null);
}
