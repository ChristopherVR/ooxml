/**
 * Charts, drawing shapes and text boxes for the Word PDF export. Their SVG (from the shared chart
 * and shape painters, the drawing Print Layout shows) goes through the shared SVG to PDF converter,
 * so they stay vector paths with real text; a text box's lines are written as the page engine laid
 * them out, clipped to its text rectangle and turned for vertical text, as `print-text-box.ts`
 * draws them. Reading the SVG needs a `DOMParser`: without one (a DOM-less run) charts and shapes
 * are reported as left out, and text box lines are still written.
 */
import type { LayoutBlockBox } from 'ooxml-core/docx/layout';
import { collectSvgImages, drawSvgSubtree } from '../../export/pdf-svg';
import { PDF_SVG_NOTES } from '../../export/pdf-svg-notes';
import { PdfFallback } from '../../export/pdf-svg-style';
import { viewBoxMatrix } from '../../export/pdf-svg-viewbox';
import { PdfVectorPage, type PdfImageResource } from '../../export/pdf-vector';
import { paintBlock } from './pdf-blocks';
import { blockPictures } from './pdf-objects';
import type { DocxPdfPicture, DocxPdfPictures } from './pdf-pictures';
import type { DocxPdfReportBuilder } from './pdf-report';
import type { DocxPdfText } from './pdf-text';

interface PreparedSvg {
	svg: Element;
	images: Map<string, PdfImageResource>;
}

const DATA_URL = /^data:([\w.+-]+\/[\w.+-]+)?((?:;[\w-]+=[^;,]*)*)(;base64)?,(.*)$/s;
const WORD_ART_NOTE =
	'WordArt warps, gradients, outlines and shadows are written to PDF as plain text in the text box.';

/** The first family of a CSS font stack, unquoted: the name supplied fonts are matched by. */
function firstFamily(stack: string): string {
	return (stack.split(',')[0] ?? '').trim().replace(/^(['"])(.*)\1$/, '$2');
}

/** Text in a box's own colour where its runs set none (a shape's `textColor`). */
export function withTextColor(blocks: readonly LayoutBlockBox[], color: string | undefined) {
	if (!color) return blocks;
	const paint = (block: LayoutBlockBox): LayoutBlockBox =>
		block.kind === 'paragraph'
			? {
					...block,
					lines: block.lines.map((line) => ({
						...line,
						fragments: line.fragments.map((fragment) =>
							fragment.color ? fragment : { ...fragment, color },
						),
					})),
				}
			: {
					...block,
					rows: block.rows.map((row) => ({
						...row,
						cells: row.cells.map((cell) => cell.map((p) => paint(p) as typeof p)),
					})),
				};
	return blocks.map(paint);
}

export class DocxPdfDrawings {
	/** Parsed SVG by its markup, or why it cannot be written as vectors. */
	readonly #svgs = new Map<string, PreparedSvg | string>();
	constructor(
		readonly pictures: DocxPdfPictures,
		readonly text: DocxPdfText,
		readonly report: DocxPdfReportBuilder,
	) {}

	/** Parses the object's SVG and loads its pictures, so painting stays synchronous. */
	async prepare(object: DocxPdfPicture): Promise<void> {
		const markup = object.chart?.svg;
		if (markup && !this.#svgs.has(markup)) this.#svgs.set(markup, await this.#parse(markup));
		if (object.textBoxLayout)
			await this.pictures.prepare(blockPictures(object.textBoxLayout.blocks));
	}

	async #parse(markup: string): Promise<PreparedSvg | string> {
		if (typeof DOMParser === 'undefined')
			return 'Charts and drawing shapes need a DOMParser to be exported to PDF; this run has none.';
		const svg = new DOMParser().parseFromString(markup, 'image/svg+xml').documentElement;
		if (svg?.localName !== 'svg') return 'The drawing could not be read.';
		for (const node of [svg, ...Array.from(svg.querySelectorAll('[font-family]'))])
			if (node.hasAttribute('font-family'))
				node.setAttribute('font-family', firstFamily(node.getAttribute('font-family')!));
		const images = new Map<string, PdfImageResource>();
		for (const href of collectSvgImages(svg)) {
			const image = await this.#dataImage(href);
			if (image) images.set(href, image);
		}
		try {
			// A trial run: the converter throws before anything reaches the real page.
			drawSvgSubtree(new PdfVectorPage(10, 10), svg, images, this.text.fonts);
		} catch (error) {
			if (error instanceof PdfFallback)
				return `The drawing has content the PDF export cannot write as vectors (${error.message}).`;
			throw error;
		}
		return { svg, images };
	}

	/** A picture embedded in the SVG as a data URL; nothing else is ever fetched. */
	async #dataImage(href: string): Promise<PdfImageResource | undefined> {
		const match = DATA_URL.exec(href);
		if (!match?.[3] || typeof atob === 'undefined') return undefined;
		try {
			const binary = atob(match[4]!.trim());
			const bytes = new Uint8Array(binary.length);
			for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
			return (await this.pictures.decode(bytes, match[1] ?? '')) ?? undefined;
		} catch {
			return undefined;
		}
	}

	/** Draws a chart or shape (and its text box) into its box at (x, y). */
	draw(page: PdfVectorPage, object: DocxPdfPicture, x: number, y: number): void {
		const chart = object.chart;
		if (!chart) return;
		const kind = chart.shape ? 'shape' : 'chart';
		const prepared = chart.svg ? this.#svgs.get(chart.svg) : undefined;
		if (!chart.svg)
			this.report.omitted(kind, `${chart.label} has no drawing to export (only its frame).`);
		else if (!prepared || typeof prepared === 'string')
			this.report.omitted(kind, prepared ?? 'The drawing was not prepared for the export.');
		else this.#paintSvg(page, prepared, object, x, y, !!chart.shape);
		if (object.textBoxLayout) this.#paintTextBox(page, object, x, y);
		else if (object.textBox)
			this.report.omitted('textBox', 'A text box the page engine did not lay out was left out.');
	}

	#paintSvg(
		page: PdfVectorPage,
		{ svg, images }: PreparedSvg,
		object: DocxPdfPicture,
		x: number,
		y: number,
		shape: boolean,
	): void {
		// A shape is drawn at its SVG's own size and its outline may pass its box; a chart is
		// stretched to its box and clipped to it, as Print Layout shows them.
		const size = (name: 'width' | 'height', fallback: number) =>
			shape ? Number(svg.getAttribute(name)) || fallback : fallback;
		page.save();
		if (!shape) {
			page.rect(x, y, object.widthPx, object.heightPx);
			page.op('W n');
		}
		page.concat([1, 0, 0, 1, x, y]);
		page.concat(viewBoxMatrix(svg, size('width', object.widthPx), size('height', object.heightPx)));
		for (const code of drawSvgSubtree(page, svg, images, this.text.fonts))
			this.report.note(PDF_SVG_NOTES[code]);
		page.restore();
	}

	#paintTextBox(page: PdfVectorPage, object: DocxPdfPicture, x: number, y: number): void {
		const layout = object.textBoxLayout!;
		const { frame, rotation } = layout;
		const length = rotation ? frame.heightPx : frame.widthPx;
		const depth = rotation ? frame.widthPx : frame.heightPx;
		page.save();
		page.rect(x + frame.xPx, y + frame.yPx, frame.widthPx, frame.heightPx);
		page.op('W n');
		page.concat([1, 0, 0, 1, x + frame.xPx, y + frame.yPx]);
		// Turned about the frame's top-left: `vert` puts the first line at the right, `vert270` left.
		if (rotation === 90) page.concat([0, 1, -1, 0, depth, 0]);
		if (rotation === 270) page.concat([0, -1, 1, 0, 0, length]);
		const painter = { page, text: this.text, pictures: this.pictures };
		for (const block of withTextColor(layout.blocks, object.chart?.textColor))
			paintBlock(painter, block, 0, 0);
		page.restore();
		const wordArt = object.textBox?.wordArt;
		if (wordArt && (wordArt.warp || wordArt.gradient || wordArt.outline || wordArt.shadow))
			this.report.note(WORD_ART_NOTE);
	}
}
