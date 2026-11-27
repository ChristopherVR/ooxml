/**
 * Every picture of the deck by name, drawn once: the hero of a case study appears on the index,
 * its opener and its detail slide, and the package stores it a single time.
 */
import { bakehouseBoard, bakehousePackaging } from './art-bakehouse';
import { gardenPictograms, gardenSigns } from './art-garden';
import { outfittersCover, outfittersSpread } from './art-outfitters';
import { coverComposition, portrait } from './art-studio';

const DRAW = {
	cover: coverComposition,
	hero1: bakehouseBoard,
	hero2: gardenSigns,
	hero3: outfittersCover,
	detail1: bakehousePackaging,
	detail2: gardenPictograms,
	detail3: outfittersSpread,
	portrait1: () => portrait(0),
	portrait2: () => portrait(1),
	portrait3: () => portrait(2),
	portrait4: () => portrait(3),
	portrait5: () => portrait(4),
} as const;

export type PictureName = keyof typeof DRAW;

const drawn = new Map<PictureName, Uint8Array>();

export function png(name: PictureName): Uint8Array {
	let bytes = drawn.get(name);
	if (!bytes) {
		bytes = DRAW[name]();
		drawn.set(name, bytes);
	}
	return bytes;
}

/** What each picture shows, for its alternative text. */
export const ALT: Record<PictureName, string> = {
	cover:
		'A composition of a sage arch, an ink-blue quarter disc, a vermilion disc and a thimble, joined by one thread',
	hero1:
		'Identity board for Brindlemoor Bakehouse: a wheat ear on terracotta, a round stamp, an awning stripe and a loaf pictogram',
	hero2:
		'Three wayfinding totems in the Larkspur Heights Community Garden, each with a coloured head, a pictogram and an arrow',
	hero3:
		'Cover of the Kestrelline Outfitters field guide: six ridgelines at dusk with contour rings and a dotted route',
	detail1: 'Brindlemoor packaging: a paper loaf bag, a pastry box, a cup and a round label',
	detail2:
		'Eight garden pictograms on coloured tiles: leaf, water, sun, arrow, seedling, tree, bench, gate',
	detail3: 'An open spread of the field guide with a contour map, a route and an elevation profile',
	portrait1: 'Abstract portrait in flat colour',
	portrait2: 'Abstract portrait in flat colour',
	portrait3: 'Abstract portrait in flat colour',
	portrait4: 'Abstract portrait in flat colour',
	portrait5: 'Abstract portrait in flat colour',
};
