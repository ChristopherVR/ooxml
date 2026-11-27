/**
 * The three case studies: an opener with one large picture, then a detail slide. The picture, the
 * case number and a short vermilion rule carry the same `!!` names on both, so Morph moves them.
 */
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import type { Deck } from '../kit/deck';
import { pt } from '../kit/geometry';
import type { Slide } from '../kit/slide';
import { CASES, type CaseStudy } from './content';
import { ALT, png, type PictureName } from './pictures';
import { C, G, T, col, figure, headline, kicker, label, numbered, rule, span } from './style';
import { LAYOUT } from './template';

const hero = (at: number) => `hero${at + 1}` as PictureName;
const detail = (at: number) => `detail${at + 1}` as PictureName;

/** The short accent rule that travels with the case number. */
function accent(slide: Slide, at: number, x: number, y: number): PptxElement {
	return slide.shape(`!!case${at + 1} rule`, 'rect', {
		x,
		y,
		width: 48,
		height: 3,
		fill: { type: 'solid', color: C.vermilion },
		stroke: { color: 'transparent' },
	});
}

function bigNumber(slide: Slide, at: number, study: CaseStudy): PptxElement {
	return headline(slide, `!!case${at + 1} number`, study.number, {
		x: G.left - 4,
		y: 48,
		width: 240,
		height: 110,
		size: 84,
		color: C.vermilion,
	});
}

function lightOpener(deck: Deck, at: number, study: CaseStudy): void {
	const slide = deck.slide(LAYOUT.opener);
	const x = col(4);
	const width = G.right - x;
	const height = width / 1.6;
	slide.picture(`!!case${at + 1} image`, png(hero(at)), {
		x,
		y: G.top,
		width,
		height,
		altText: ALT[hero(at)],
	});
	bigNumber(slide, at, study);
	accent(slide, at, G.left, 352);
	kicker(slide, 'Case kicker', `${study.discipline}, ${study.year}`, G.left, 370, span(3));
	slide.placeholder('title', study.client);
	const caption = label(slide, 'Case caption', study.caption, {
		x,
		y: G.top + height + 20,
		width: span(6),
		height: 54,
		fontSize: pt(T.caption + 1),
		color: C.soft,
		lineSpacing: 1.35,
	});
	numbered(slide);
	slide.animate(caption, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 700 });
	slide.notes(study.notes.opener);
}

function darkOpener(deck: Deck, at: number, study: CaseStudy): void {
	const slide = deck.slide(LAYOUT.dark);
	// The picture is 16:10; it overshoots the slide top and bottom so that it bleeds on all sides.
	slide.picture(`!!case${at + 1} image`, png(hero(at)), {
		x: 0,
		y: -40,
		width: 1280,
		height: 800,
		altText: ALT[hero(at)],
	});
	const scrim = slide.shape('Scrim', 'rect', {
		x: 0,
		y: 0,
		width: 820,
		height: 720,
		fill: {
			type: 'gradient',
			angle: 0,
			stops: [
				{ color: C.night, position: 0, opacity: 0.92 },
				{ color: C.night, position: 55, opacity: 0.7 },
				{ color: C.night, position: 100, opacity: 0 },
			],
		},
		stroke: { color: 'transparent' },
	});
	bigNumber(slide, at, study);
	accent(slide, at, G.left, 318);
	kicker(slide, 'Case kicker', `${study.discipline}, ${study.year}`, G.left, 336, span(4), C.clay);
	slide.placeholder('title', study.client.split(' '));
	const caption = label(slide, 'Case caption', study.caption, {
		x: G.left,
		y: 540,
		width: span(5),
		height: 60,
		fontSize: pt(T.caption + 1),
		color: C.clay,
		lineSpacing: 1.35,
	});
	numbered(slide);
	slide.animate(caption, { entrance: 'fadeIn', trigger: 'afterPrevious', durationMs: 700 });
	// On a click the scrim leaves, and the cover is seen as printed.
	slide.animate(scrim, { exit: 'fadeOut', trigger: 'onClick', durationMs: 900 });
	slide.animate(caption, { exit: 'fadeOut', trigger: 'withPrevious', durationMs: 600 });
	slide.notes(study.notes.opener);
}

function detailSlide(deck: Deck, at: number, study: CaseStudy): Slide {
	const slide = deck.slide(LAYOUT.detail);
	const right = col(5);
	slide.picture(`!!case${at + 1} image`, png(hero(at)), {
		x: G.left,
		y: 96,
		width: 400,
		height: 250,
		altText: ALT[hero(at)],
	});
	slide.picture('Detail picture', png(detail(at)), {
		x: G.left,
		y: 362,
		width: 400,
		height: 250,
		altText: ALT[detail(at)],
	});
	headline(slide, `!!case${at + 1} number`, study.number, {
		x: right,
		y: 56,
		width: 40,
		height: 22,
		size: 15,
		color: C.vermilion,
	});
	kicker(slide, 'Case kicker', `${study.client}, ${study.year}`, right + 40, 62, span(6));
	accent(slide, at, right, 84);
	slide.placeholder('title', study.headline);
	slide.placeholder('body', study.brief, 1);
	kicker(slide, 'Made kicker', 'What we made', right, 280, span(4));
	label(slide, 'Made list', study.made, {
		x: right,
		y: 304,
		width: span(7),
		height: 96,
		fontSize: pt(T.body),
		lineSpacing: 1.3,
		spaceAfter: 7,
	});
	rule(slide, 'Figures rule', right, 424, span(7));
	const width = (span(7) - 2 * G.gutter) / 3;
	study.figures.forEach(([value, caption], index) =>
		figure(
			slide,
			`Figure ${index + 1}`,
			right + index * (width + G.gutter),
			446,
			width,
			value,
			caption,
		),
	);
	numbered(slide);
	slide.notes(study.notes.detail);
	return slide;
}

export function caseSlides(deck: Deck): void {
	const [bakehouse, garden, outfitters] = CASES as [CaseStudy, CaseStudy, CaseStudy];

	lightOpener(deck, 0, bakehouse);
	deck.slides[deck.slides.length - 1]!.transition({ type: 'morph', duration: 1400 });
	// A paragraph build written by hand: one line per click, then the figures together.
	const first = detailSlide(deck, 0, bakehouse);
	first.effects.push(
		{ target: 'Made list', paragraph: 0, kind: 'rise' },
		{ target: 'Made list', paragraph: 1, kind: 'rise' },
		{ target: 'Made list', paragraph: 2, kind: 'rise' },
		...[1, 2, 3].flatMap((n) => [
			{ target: `Figure ${n} value`, kind: 'fade', start: n === 1 ? 'click' : 'after' } as const,
			{ target: `Figure ${n} caption`, kind: 'fade', start: 'with' } as const,
		]),
	);
	first.transition({ type: 'morph', duration: 1200 });

	lightOpener(deck, 1, garden);
	deck.slides[deck.slides.length - 1]!.transition({ type: 'cover', direction: 'l', duration: 800 });
	const second = detailSlide(deck, 1, garden);
	const find = (name: string) => second.model.elements.find((element) => element.name === name)!;
	second.animate(find('Made list'), {
		entrance: 'floatIn',
		trigger: 'afterPrevious',
		durationMs: 700,
	});
	// The pictogram sheet pulses once when the list has arrived.
	second.animate(find('Detail picture'), {
		emphasis: 'pulse',
		trigger: 'afterPrevious',
		durationMs: 800,
	});
	[1, 2, 3].forEach((n) => {
		second.animate(find(`Figure ${n} value`), {
			entrance: 'fadeIn',
			trigger: n === 1 ? 'onClick' : 'afterPrevious',
			durationMs: 450,
		});
		second.animate(find(`Figure ${n} caption`), {
			entrance: 'fadeIn',
			trigger: 'withPrevious',
			durationMs: 450,
		});
	});
	second.transition({ type: 'morph', duration: 1200 });

	darkOpener(deck, 2, outfitters);
	deck.slides[deck.slides.length - 1]!.transition({ type: 'split', duration: 900 });
	const third = detailSlide(deck, 2, outfitters);
	third.effects.push(
		{ target: 'Made list', paragraph: 0, kind: 'wipe', duration: 600 },
		{ target: 'Made list', paragraph: 1, kind: 'wipe', duration: 600 },
		{ target: 'Made list', paragraph: 2, kind: 'wipe', duration: 600 },
		{ target: 'Figures rule', kind: 'wipe', start: 'click', duration: 500 },
		...[1, 2, 3].flatMap((n) => [
			{ target: `Figure ${n} value`, kind: 'rise', start: 'after', duration: 400 } as const,
			{ target: `Figure ${n} caption`, kind: 'rise', start: 'with', duration: 400 } as const,
		]),
	);
	third.transition({ type: 'morph', duration: 1400 });
}
