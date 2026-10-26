import { vi } from 'vitest';

/** A one-face embedded font list, as one loaded document would carry. */
export function embeddedFontList(name = 'OutlineReleaseTest') {
	return [{ name, dataUrl: '', bold: false, italic: false }];
}

/**
 * Give every `Handler#load` result a font list of its own (the sample decks
 * embed none), returned in load order.
 */
export function embedFontsOnLoad(Handler) {
	const load = Handler.prototype.load;
	const lists = [];
	vi.spyOn(Handler.prototype, 'load').mockImplementation(async function (...args) {
		const parsed = await load.apply(this, args);
		const list = embeddedFontList();
		lists.push(list);
		parsed.embeddedFonts = list;
		return parsed;
	});
	return lists;
}

/** The font lists a glyph-outline cache still holds undecoded candidates of. */
export function retainedFontLists(cache) {
	return new Set([...cache.pendingEmbedded].map((candidate) => candidate.list));
}
