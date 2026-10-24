/**
 * The `a:fmtScheme` elements a handler has written into its theme parts or over them, so Undo and
 * Redo of a Design > Effects pick can put each part back as it was (`restoreThemeFormatScheme`).
 *
 * A parsed format scheme is lossy, so an undo record cannot write its scheme back itself: it names
 * one (`formatSchemeKey`) and this hands back the element. A pick writes into every part a master
 * uses, and the record holds the deck's scheme only (the first master's part), so the parts are
 * also kept by the deck's scheme and the parts the masters used at the time: putting a deck scheme
 * back over the same parts puts each part back. The same deck scheme over other parts (the masters
 * moved to another theme part in between) is another entry.
 */
import type { PptxThemeFormatScheme, XmlObject } from '../../types';
import { formatSchemeKey } from '../../utils/theme-effect-schemes';

export class HeldFormatSchemes {
	private readonly elements = new Map<string, XmlObject>();
	private readonly partsByDeckScheme = new Map<string, ReadonlyMap<string, XmlObject>>();

	/** Keep `element`, which `scheme` was parsed from. */
	hold(scheme: PptxThemeFormatScheme, element: XmlObject): void {
		this.elements.set(formatSchemeKey(scheme), structuredClone(element));
	}

	/** The element `scheme` was parsed from, if one was held. */
	element(scheme: PptxThemeFormatScheme): XmlObject | undefined {
		return this.elements.get(formatSchemeKey(scheme));
	}

	/** What each part in `parts` held while the deck's scheme was `deckScheme`. */
	holdParts(deckScheme: PptxThemeFormatScheme, parts: ReadonlyMap<string, XmlObject>): void {
		if (parts.size > 0) {
			this.partsByDeckScheme.set(partsKey(deckScheme, [...parts.keys()]), new Map(parts));
		}
	}

	/**
	 * The element each of `themePaths` held the last time the deck's scheme was `deckScheme` over
	 * exactly those parts, if known.
	 */
	parts(
		deckScheme: PptxThemeFormatScheme,
		themePaths: readonly string[],
	): ReadonlyMap<string, XmlObject> | undefined {
		return this.partsByDeckScheme.get(partsKey(deckScheme, themePaths));
	}
}

function partsKey(deckScheme: PptxThemeFormatScheme, themePaths: readonly string[]): string {
	return JSON.stringify([formatSchemeKey(deckScheme), [...new Set(themePaths)].sort()]);
}
