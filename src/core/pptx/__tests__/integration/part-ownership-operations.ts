import { duplicateElement, duplicateSlide } from '../../core/builders/sdk/slide-operations';
import {
	duplicateLayout,
	duplicateSlideMaster,
	insertSlideMaster,
} from '../../core/builders/sdk/master-layout-crud';
import type { PptxHandler } from '../../core/PptxHandler';
import type { PptxData, PptxElement, PptxSlide } from '../../core/types';
import { duplicateElementOnSlide } from '../../editor/render/editor-mutations';
import { createBlankSlide } from '../../editor/render/slide-operations';

export interface Deck {
	handler: PptxHandler;
	data: PptxData;
}

const PART_BEARING = new Set<PptxElement['type']>([
	'chart',
	'smartArt',
	'ole',
	'media',
	'picture',
	'image',
	'group',
]);

/** Elements that show package parts: what every copy below copies. */
const partBearing = (slide: PptxSlide): PptxElement[] =>
	slide.elements.filter(
		(element) => PART_BEARING.has(element.type) && element.rawXml !== undefined,
	);

/**
 * What a clipboard round trip gives a binding to paste (`element-clipboard`): a new id, and no
 * native shape id, which the save allocates.
 */
function clipboardCopy(element: PptxElement, id: string): PptxElement {
	const copy = { ...(JSON.parse(JSON.stringify(element)) as PptxElement), id };
	delete copy.shapeId;
	return copy;
}

type Operation = (deck: Deck) => Promise<Deck | void>;

export const OPERATIONS = {
	'duplicate element (sdk)': async ({ data }) => {
		for (const slide of data.slides)
			for (const element of partBearing(slide)) slide.elements.push(duplicateElement(element));
	},
	'duplicate element (editor)': async (deck) => {
		deck.data.slides.forEach((slide, index) => {
			for (const element of partBearing(slide)) {
				const next = duplicateElementOnSlide(deck.data.slides, index, element.id);
				if (next) deck.data.slides = next.slides;
			}
		});
	},
	'shallow element copy': async ({ data }) => {
		for (const slide of data.slides)
			for (const element of partBearing(slide))
				slide.elements.push({ ...element, id: `${element.id}-shallow` });
	},
	'paste on the next slide': async ({ data }) => {
		const copies = data.slides.map((slide) =>
			partBearing(slide).map((element) => clipboardCopy(element, `${element.id}-pasted`)),
		);
		data.slides.forEach((slide, index) => {
			const next = data.slides[(index + 1) % data.slides.length]!;
			next.elements.push(...copies[index]!);
			void slide;
		});
	},
	'paste on a new blank slide': async ({ data }) => {
		const blank = createBlankSlide(data.slides.length + 1);
		blank.layoutPath = data.slides[0]?.layoutPath;
		for (const slide of data.slides)
			for (const element of partBearing(slide))
				blank.elements.push(clipboardCopy(element, `${element.id}-blank`));
		data.slides.push(blank);
	},
	'move to another slide, delete the source': async ({ data }) => {
		const source = data.slides.find((slide) => partBearing(slide).length > 0);
		if (!source || data.slides.length < 2) return;
		const destination = data.slides.find((slide) => slide !== source)!;
		destination.elements.push(...partBearing(source).map((e) => clipboardCopy(e, `${e.id}-m`)));
		data.slides = data.slides.filter((slide) => slide !== source);
	},
	'duplicate slide (sdk)': async ({ data }) => {
		for (const slide of [...data.slides])
			if (partBearing(slide).length > 0)
				data.slides.push(duplicateSlide(slide, data.slides.length + 1));
	},
	'duplicate slide then its elements': async ({ data }) => {
		const source = data.slides.find((slide) => partBearing(slide).length > 0);
		if (!source) return;
		const copy = duplicateSlide(source, data.slides.length + 1);
		copy.elements.push(...partBearing(copy).map(duplicateElement));
		data.slides.push(copy);
	},
	'duplicate master': async ({ handler, data }) => {
		const result = await duplicateSlideMaster(handler, data, data.slideMasters![0]!.path);
		if (result.ok) return result;
	},
	'insert master': async ({ handler, data }) => {
		const result = await insertSlideMaster(handler, data);
		if (result.ok) return result;
	},
	'duplicate every layout': async (deck) => {
		let current = deck;
		for (const layout of [...(deck.data.slideMasters?.[0]?.layouts ?? [])]) {
			const result = await duplicateLayout(current.handler, current.data, layout.path);
			if (result.ok) {
				if (current !== deck) current.handler.dispose();
				current = result;
			}
		}
		return current;
	},
} satisfies Record<string, Operation>;
