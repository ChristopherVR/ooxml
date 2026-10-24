/**
 * One print job plan for every binding's Print dialog.
 *
 * The five bindings used to re-decide, each in its own driver, which slides a
 * confirmed Print prints (range, current slide, hidden slides), which path
 * renders them (text outline, core SVG slides, captured slide images), whether
 * the print window has to open before the first capture, and what the print
 * window is called. They drifted: one never reached its vector path, one
 * captured every slide of the deck to print one, four titled the window in
 * English only. {@link planPrintJob} makes those decisions once and returns a
 * framework-neutral descriptor; the `build*PrintJobDocument` functions turn a
 * plan plus what the binding captured into the printable document. A binding
 * keeps only its capture (SVG export, stage rasterisation) and its opener
 * (window, iframe or host callback).
 *
 * Pure: no DOM access beyond the DOMPurify pass the shared document builders
 * already run, no `window.open`, no `print()`.
 */
import type { PptxHandoutMaster, PptxSlide } from 'ooxml-core/pptx';

import {
	buildHandoutsHtml,
	buildNotesHtml,
	buildOutlineHtml,
	buildPrintHtmlDocument,
	buildSlidesHtml,
	computeColorFilter,
	computeSlideIndices,
	filterHiddenSlideIndices,
	validatePrintSettings,
} from './print-document';
import type { PrintOrientation, PrintSettings, PrintSlideRange, PrintWhat } from './print-document';
import { buildPrintDocument, sanitizePixelDimension } from './svg-print';

/**
 * The confirmed print settings: any subset, an `undefined` field meaning its
 * default (`DEFAULT_PRINT_SETTINGS`).
 */
export type PrintJobSettings = {
	readonly [K in keyof PrintSettings]?: PrintSettings[K] | undefined;
};

/** What the binding knows about the deck when the user confirms Print. */
export interface PrintJobDeck {
	/**
	 * The complete slides, in deck order, with their layout and master elements
	 * merged back in (what `buildSaveSlides` returns): the vector path exports
	 * these, and the outline and notes pages read their text.
	 */
	readonly slides: readonly PptxSlide[];
	/** Zero-based index of the slide on screen; the `current` range prints it. */
	readonly activeSlideIndex: number;
	/** Options > Advanced > "Print hidden slides". Defaults to `false`, as in PowerPoint. */
	readonly includeHiddenSlides?: boolean | undefined;
}

/** Why a print job has nothing to print. */
export type PrintJobRefusalReason =
	/** `printWhat` or `slideRange` is not a value the print dialog offers. */
	| 'invalid-settings'
	/** The range selects no slide that may print (empty deck, or only hidden slides). */
	| 'empty-selection';

/** A confirmed Print that must not open a print window. */
export interface PrintJobRefusal {
	readonly ok: false;
	readonly reason: PrintJobRefusalReason;
}

/**
 * How the plan's pages are produced: `outline` is text only, `vector` exports
 * each planned slide through the core SVG exporter, `raster` captures each
 * planned slide as an image (full-page slides, notes pages or handouts).
 */
export type PrintJobOutput = 'outline' | 'vector' | 'raster';

/** The print window title as a translation key; the binding translates it. */
export interface PrintJobTitle {
	readonly key: string;
	/** Interpolation values; only the handouts title has one (`{{count}}`). */
	readonly params?: { readonly count: number };
}

/** A print job to run. */
export interface PrintJobPlan {
	readonly ok: true;
	/** The validated settings (custom range clamped, orientation resolved). */
	readonly settings: PrintSettings;
	/** What the pages show. */
	readonly printWhat: PrintWhat;
	/** Zero-based slides to print, in print order; never empty. */
	readonly slideIndices: readonly number[];
	readonly output: PrintJobOutput;
	/**
	 * `true` when the document can only be assembled after awaiting slide
	 * captures. A `window.open` based opener must then open its window before
	 * the first await (`openPendingPrintWindow`), or the browser blocks it as a
	 * popup; it can open the finished document in one call otherwise.
	 */
	readonly openWindowFirst: boolean;
	/** Page orientation: the chosen one for full-page slides, portrait otherwise. */
	readonly orientation: PrintOrientation;
	readonly title: PrintJobTitle;
}

/** The outcome of {@link planPrintJob}. */
export type PrintJob = PrintJobPlan | PrintJobRefusal;

/** One captured slide image for a raster plan. */
export interface PrintJobCapture {
	/** Zero-based index of the captured slide (one of the plan's `slideIndices`). */
	readonly slideIndex: number;
	/** The slide image as a `data:image/...` URL. */
	readonly dataUrl: string;
}

/** Slide size in CSS pixels, for the vector document. */
export interface PrintJobSlideSize {
	readonly width: number;
	readonly height: number;
}

const PRINT_WHAT: readonly PrintWhat[] = ['slides', 'handouts', 'notes', 'outline'];
const SLIDE_RANGES: readonly PrintSlideRange[] = ['all', 'current', 'custom'];

/** The print window title of a print mode, as the translation keys every locale already has. */
function printJobTitle(settings: PrintSettings): PrintJobTitle {
	if (settings.printWhat === 'notes') {
		return { key: 'pptx.print.notesPages' };
	}
	if (settings.printWhat === 'outline') {
		return { key: 'pptx.print.outline' };
	}
	if (settings.printWhat === 'handouts') {
		return { key: 'pptx.print.handoutPerPageTitle', params: { count: settings.slidesPerPage } };
	}
	return { key: 'pptx.sections.slides' };
}

/** Drop `undefined` entries so an explicitly unset field takes its default. */
function definedSettings(settings: PrintJobSettings): Partial<PrintSettings> {
	return Object.fromEntries(
		Object.entries(settings).filter(([, value]) => value !== undefined),
	) as Partial<PrintSettings>;
}

/**
 * Decide what a confirmed Print prints. `settings` is validated against the
 * deck (custom range clamped, slides per page normalised). Full-page slides
 * print as vector SVG (fall back with {@link printJobRasterFallback} if the
 * export fails), notes pages and handouts as captured slide images, the
 * outline as text. Hidden slides are left out unless
 * `deck.includeHiddenSlides` is set.
 */
export function planPrintJob(settings: PrintJobSettings, deck: PrintJobDeck): PrintJob {
	const slideCount = deck.slides.length;
	const validated = validatePrintSettings(definedSettings(settings), slideCount);
	if (!PRINT_WHAT.includes(validated.printWhat) || !SLIDE_RANGES.includes(validated.slideRange)) {
		return { ok: false, reason: 'invalid-settings' };
	}
	const inDeck = computeSlideIndices(
		validated.slideRange,
		deck.activeSlideIndex,
		slideCount,
		validated.customRangeFrom,
		validated.customRangeTo,
	).filter((index) => Number.isInteger(index) && index >= 0 && index < slideCount);
	const slideIndices = filterHiddenSlideIndices(
		inDeck,
		deck.slides,
		deck.includeHiddenSlides ?? false,
	);
	if (slideIndices.length === 0) {
		return { ok: false, reason: 'empty-selection' };
	}
	let output: PrintJobOutput = 'raster';
	if (validated.printWhat === 'outline') {
		output = 'outline';
	} else if (validated.printWhat === 'slides') {
		output = 'vector';
	}
	return {
		ok: true,
		settings: validated,
		printWhat: validated.printWhat,
		slideIndices,
		output,
		openWindowFirst: output === 'raster',
		orientation: validated.orientation,
		title: printJobTitle(validated),
	};
}

/**
 * The same job captured as slide images instead: what a binding runs when
 * exporting a `vector` plan's slides failed. Any other plan is returned as is.
 */
export function printJobRasterFallback(plan: PrintJobPlan): PrintJobPlan {
	if (plan.output !== 'vector') {
		return plan;
	}
	return { ...plan, output: 'raster', openWindowFirst: true };
}

/**
 * Draw the Frame Slides border inside a slide SVG, before its closing tag. The
 * last `</svg>` is the root's own: a slide SVG may nest others (charts).
 */
export function frameSlideSvg(svg: string, width: number, height: number): string {
	const end = svg.lastIndexOf('</svg>');
	if (end < 0) {
		return svg;
	}
	const frameWidth = Math.max(0, sanitizePixelDimension(width) - 2);
	const frameHeight = Math.max(0, sanitizePixelDimension(height) - 2);
	const frame = `<rect x="1" y="1" width="${frameWidth}" height="${frameHeight}" fill="none" stroke="#000" stroke-width="2" />`;
	return `${svg.slice(0, end)}${frame}${svg.slice(end)}`;
}

/** The printable document of an `outline` plan. `title` is the translated `plan.title`. */
export function buildOutlinePrintJobDocument(
	plan: PrintJobPlan,
	title: string,
	slides: readonly PptxSlide[],
): string {
	const outline = buildOutlineHtml([...plan.slideIndices], slides);
	return buildPrintHtmlDocument({
		title,
		bodyHtml: `<div class="outline-page">${outline}</div>`,
		orientation: plan.orientation,
		colorFilter: computeColorFilter(plan.settings.colorMode),
		frameSlides: plan.settings.frameSlides,
		scaleToFit: plan.settings.scaleToFit ?? true,
	});
}

/**
 * The printable document of a `vector` plan, from the SVG of each planned
 * slide in plan order. Frame Slides draws a border inside each slide.
 */
export function buildVectorPrintJobDocument(
	plan: PrintJobPlan,
	title: string,
	svgs: readonly string[],
	size: PrintJobSlideSize,
): string {
	const pages = plan.settings.frameSlides
		? svgs.map((svg) => frameSlideSvg(svg, size.width, size.height))
		: [...svgs];
	return buildPrintDocument(pages, size.width, size.height, {
		title,
		orientation: plan.orientation,
		colorFilter: computeColorFilter(plan.settings.colorMode),
		scaleToFit: plan.settings.scaleToFit ?? true,
	});
}

/**
 * The printable document of a `raster` plan: full-page slides, notes pages or
 * handouts, from the slide images the binding captured (a slide whose capture
 * failed is simply absent). `slides` supplies the notes text, `handoutMaster`
 * the handout page chrome.
 */
export function buildRasterPrintJobDocument(
	plan: PrintJobPlan,
	title: string,
	captures: readonly PrintJobCapture[],
	slides: readonly PptxSlide[],
	handoutMaster?: PptxHandoutMaster,
): string {
	const images = captures.map((capture) => capture.dataUrl);
	const indices = captures.map((capture) => capture.slideIndex);
	let bodyHtml = buildSlidesHtml(images, indices);
	if (plan.printWhat === 'notes') {
		bodyHtml = buildNotesHtml(images, indices, slides);
	} else if (plan.printWhat === 'handouts') {
		bodyHtml = buildHandoutsHtml(images, indices, plan.settings.slidesPerPage, handoutMaster);
	}
	return buildPrintHtmlDocument({
		title,
		bodyHtml,
		orientation: plan.orientation,
		colorFilter: computeColorFilter(plan.settings.colorMode),
		frameSlides: plan.settings.frameSlides,
		scaleToFit: plan.settings.scaleToFit ?? true,
	});
}
