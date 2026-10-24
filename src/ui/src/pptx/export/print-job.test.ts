// @vitest-environment jsdom
//
// jsdom gives DOMPurify a window, so the document builders below sanitise for
// real (outside a DOM they fail closed with an empty body).
import type { PptxHandoutMaster, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	buildOutlinePrintJobDocument,
	buildRasterPrintJobDocument,
	buildVectorPrintJobDocument,
	frameSlideSvg,
	planPrintJob,
	printJobRasterFallback,
} from './print-job';
import type { PrintJob, PrintJobDeck, PrintJobPlan } from './print-job';

function makeSlide(index: number, overrides: Partial<PptxSlide> = {}): PptxSlide {
	return {
		id: `s${index}`,
		rId: `rId${index}`,
		slideNumber: index + 1,
		elements: [
			{
				id: `t${index}`,
				type: 'text',
				x: 0,
				y: 0,
				width: 100,
				height: 20,
				text: `Title ${index + 1}`,
			},
		],
		notes: `Notes ${index + 1}`,
		...overrides,
	} as PptxSlide;
}

function deck(count: number, overrides: Partial<PrintJobDeck> = {}): PrintJobDeck {
	return {
		slides: Array.from({ length: count }, (_, i) => makeSlide(i)),
		activeSlideIndex: 0,
		...overrides,
	};
}

function planned(job: PrintJob): PrintJobPlan {
	if (!job.ok) {
		throw new Error(`expected a plan, got the refusal ${job.reason}`);
	}
	return job;
}

const SVG =
	'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540"><text>slide</text></svg>';

describe('planPrintJob', () => {
	it('prints every slide of the deck as vector slides by default', () => {
		const plan = planned(planPrintJob({}, deck(3)));
		expect(plan.printWhat).toBe('slides');
		expect(plan.output).toBe('vector');
		expect(plan.slideIndices).toStrictEqual([0, 1, 2]);
		expect(plan.openWindowFirst).toBeFalsy();
		expect(plan.orientation).toBe('landscape');
		expect(plan.title).toStrictEqual({ key: 'pptx.sections.slides' });
	});

	it('keeps the chosen orientation for full-page slides', () => {
		const plan = planned(planPrintJob({ orientation: 'portrait' }, deck(1)));
		expect(plan.orientation).toBe('portrait');
	});

	it('prints only the slide on screen for the current range', () => {
		const plan = planned(planPrintJob({ slideRange: 'current' }, deck(5, { activeSlideIndex: 3 })));
		expect(plan.slideIndices).toStrictEqual([3]);
	});

	it('clamps a custom range to the deck', () => {
		const plan = planned(
			planPrintJob({ slideRange: 'custom', customRangeFrom: 2, customRangeTo: 99 }, deck(4)),
		);
		expect(plan.slideIndices).toStrictEqual([1, 2, 3]);
		expect(plan.settings.customRangeTo).toBe(4);
	});

	it('swaps an inverted custom range', () => {
		const plan = planned(
			planPrintJob({ slideRange: 'custom', customRangeFrom: 3, customRangeTo: 2 }, deck(4)),
		);
		expect(plan.slideIndices).toStrictEqual([1, 2]);
	});

	it('leaves hidden slides out unless Print hidden slides is on', () => {
		const slides = [makeSlide(0), makeSlide(1, { hidden: true }), makeSlide(2)];
		expect(planned(planPrintJob({}, { slides, activeSlideIndex: 0 })).slideIndices).toStrictEqual([
			0, 2,
		]);
		expect(
			planned(planPrintJob({}, { slides, activeSlideIndex: 0, includeHiddenSlides: true }))
				.slideIndices,
		).toStrictEqual([0, 1, 2]);
	});

	it('captures notes pages, in portrait, after opening the window', () => {
		const plan = planned(planPrintJob({ printWhat: 'notes', orientation: 'landscape' }, deck(2)));
		expect(plan.output).toBe('raster');
		expect(plan.openWindowFirst).toBeTruthy();
		expect(plan.orientation).toBe('portrait');
		expect(plan.title).toStrictEqual({ key: 'pptx.print.notesPages' });
	});

	it('captures handouts and titles them with the slides per page', () => {
		const plan = planned(planPrintJob({ printWhat: 'handouts', slidesPerPage: 4 }, deck(2)));
		expect(plan.output).toBe('raster');
		expect(plan.openWindowFirst).toBeTruthy();
		expect(plan.orientation).toBe('portrait');
		expect(plan.title).toStrictEqual({
			key: 'pptx.print.handoutPerPageTitle',
			params: { count: 4 },
		});
	});

	it('normalises an unsupported slides-per-page count', () => {
		const plan = planned(planPrintJob({ printWhat: 'handouts', slidesPerPage: 5 as 4 }, deck(1)));
		expect(plan.settings.slidesPerPage).toBe(6);
		expect(plan.title.params).toStrictEqual({ count: 6 });
	});

	it('prints the outline as text, in portrait, without opening a window first', () => {
		const plan = planned(planPrintJob({ printWhat: 'outline' }, deck(2)));
		expect(plan.output).toBe('outline');
		expect(plan.openWindowFirst).toBeFalsy();
		expect(plan.orientation).toBe('portrait');
		expect(plan.title).toStrictEqual({ key: 'pptx.print.outline' });
	});

	it('treats an explicitly undefined field as its default', () => {
		const plan = planned(planPrintJob({ printWhat: undefined, slideRange: undefined }, deck(2)));
		expect(plan.printWhat).toBe('slides');
		expect(plan.slideIndices).toStrictEqual([0, 1]);
	});

	it('refuses an empty deck in every mode, the outline included', () => {
		for (const printWhat of ['slides', 'notes', 'handouts', 'outline'] as const) {
			expect(planPrintJob({ printWhat }, deck(0))).toStrictEqual({
				ok: false,
				reason: 'empty-selection',
			});
		}
	});

	it('refuses a selection of hidden slides only', () => {
		const slides = [makeSlide(0, { hidden: true }), makeSlide(1)];
		expect(planPrintJob({ slideRange: 'current' }, { slides, activeSlideIndex: 0 })).toStrictEqual({
			ok: false,
			reason: 'empty-selection',
		});
	});

	it('refuses a current slide outside the deck', () => {
		expect(planPrintJob({ slideRange: 'current' }, deck(2, { activeSlideIndex: 5 }))).toStrictEqual(
			{ ok: false, reason: 'empty-selection' },
		);
		expect(
			planPrintJob({ slideRange: 'current' }, deck(2, { activeSlideIndex: -1 })),
		).toStrictEqual({ ok: false, reason: 'empty-selection' });
	});

	it('refuses a print mode or range the dialog does not offer', () => {
		const printWhat = 'everything' as 'slides';
		const slideRange = 'odd' as 'all';
		expect(planPrintJob({ printWhat }, deck(2))).toStrictEqual({
			ok: false,
			reason: 'invalid-settings',
		});
		expect(planPrintJob({ slideRange }, deck(2))).toStrictEqual({
			ok: false,
			reason: 'invalid-settings',
		});
	});
});

describe('printJobRasterFallback', () => {
	it('turns a vector plan into the same slides captured after opening the window', () => {
		const plan = planned(planPrintJob({ slideRange: 'current' }, deck(3, { activeSlideIndex: 1 })));
		const fallback = printJobRasterFallback(plan);
		expect(fallback.output).toBe('raster');
		expect(fallback.openWindowFirst).toBeTruthy();
		expect(fallback.slideIndices).toStrictEqual([1]);
		expect(fallback.title).toStrictEqual(plan.title);
	});

	it('returns any other plan unchanged', () => {
		const outline = planned(planPrintJob({ printWhat: 'outline' }, deck(1)));
		expect(printJobRasterFallback(outline)).toBe(outline);
	});
});

describe('frameSlideSvg', () => {
	it('draws the border inside the root svg, after any nested one', () => {
		const framed = frameSlideSvg('<svg><svg><rect /></svg><text>x</text></svg>', 100, 50);
		expect(framed).toBe(
			'<svg><svg><rect /></svg><text>x</text><rect x="1" y="1" width="98" height="48" fill="none" stroke="#000" stroke-width="2" /></svg>',
		);
	});

	it('leaves markup without a closing svg tag alone', () => {
		expect(frameSlideSvg('<g />', 100, 50)).toBe('<g />');
	});
});

describe('print job documents', () => {
	const slides = [
		makeSlide(0),
		makeSlide(1, {
			elements: [
				{ id: 'master-logo', type: 'text', x: 0, y: 0, width: 10, height: 10, text: 'Logo' },
				{ id: 't1', type: 'text', x: 0, y: 0, width: 100, height: 20, text: 'Real title' },
			] as PptxSlide['elements'],
		}),
	];

	it('builds the outline from the slide-owned titles under the translated title', () => {
		const plan = planned(planPrintJob({ printWhat: 'outline' }, { slides, activeSlideIndex: 0 }));
		const html = buildOutlinePrintJobDocument(plan, 'Gliederung', slides);
		expect(html).toContain('<title>Gliederung</title>');
		expect(html).toContain('<h2>Title 1</h2>');
		expect(html).toContain('<h2>Real title</h2>');
		expect(html).not.toContain('<h2>Logo</h2>');
		expect(html).toContain('@page { size: portrait;');
	});

	it('builds one vector page per planned slide and frames them on request', () => {
		const plan = planned(
			planPrintJob({ slideRange: 'current', frameSlides: true }, { slides, activeSlideIndex: 1 }),
		);
		const html = buildVectorPrintJobDocument(plan, 'Folien', [SVG], { width: 960, height: 540 });
		expect(html).toContain('<title>Folien</title>');
		expect(html.match(/class="print-slide-page"/gu)).toHaveLength(1);
		expect(html).toContain('stroke="#000"');
		expect(html).toContain('size: landscape;');
	});

	it('applies the colour mode to the vector document', () => {
		const plan = planned(planPrintJob({ colorMode: 'grayscale' }, { slides, activeSlideIndex: 0 }));
		const html = buildVectorPrintJobDocument(plan, 'Slides', [SVG, SVG], {
			width: 960,
			height: 540,
		});
		expect(html).toContain('filter: grayscale(1);');
		expect(html).not.toContain('stroke="#000"');
	});

	it('builds full-page raster slides from the captures', () => {
		const plan = printJobRasterFallback(
			planned(planPrintJob({ slideRange: 'current' }, { slides, activeSlideIndex: 1 })),
		);
		const html = buildRasterPrintJobDocument(
			plan,
			'Slides',
			[{ slideIndex: 1, dataUrl: 'data:image/png;base64,AAAA' }],
			slides,
		);
		expect(html).toContain('slide-page');
		expect(html).toContain('alt="Slide 2"');
	});

	it('builds notes pages with each captured slide notes text', () => {
		const plan = planned(planPrintJob({ printWhat: 'notes' }, { slides, activeSlideIndex: 0 }));
		const html = buildRasterPrintJobDocument(
			plan,
			'Notizseiten',
			[{ slideIndex: 1, dataUrl: 'data:image/png;base64,AAAA' }],
			slides,
		);
		expect(html).toContain('<title>Notizseiten</title>');
		expect(html).toContain('Notes 2');
		expect(html).not.toContain('Notes 1');
		expect(html).toContain('@page { size: portrait;');
	});

	it('builds handouts with the handout master chrome', () => {
		const handoutMaster: PptxHandoutMaster = {
			path: 'ppt/handoutMasters/handoutMaster1.xml',
			slidesPerPage: 4,
			headerFooter: { hasFooter: true },
			elements: [
				{
					id: 'ftr1',
					type: 'text',
					placeholderType: 'ftr',
					x: 0,
					y: 0,
					width: 100,
					height: 20,
					text: 'Footer text',
				} as unknown as PptxSlide['elements'][number],
			],
		};
		const plan = planned(
			planPrintJob({ printWhat: 'handouts', slidesPerPage: 4 }, { slides, activeSlideIndex: 0 }),
		);
		const html = buildRasterPrintJobDocument(
			plan,
			'Handout 4 per page',
			[
				{ slideIndex: 0, dataUrl: 'data:image/png;base64,AAAA' },
				{ slideIndex: 1, dataUrl: 'data:image/png;base64,BBBB' },
			],
			slides,
			handoutMaster,
		);
		expect(html).toContain('handout-grid');
		expect(html).toContain('Footer text');
		expect(html.match(/class="page"/gu)).toHaveLength(1);
	});
});
