/**
 * One slide of a deck, built with the pptx SDK (`SlideBuilder` and the element factories). This
 * wrapper only adds what a designed deck needs on top: every element gets a stable id and a name
 * (Morph pairs shapes by name), rich text is given as paragraphs, pictures as PNG bytes.
 */
import {
	SlideBuilder,
	createChartElement,
	createFreeformElement,
	createGroupElement,
	createImageElement,
	createShapeElement,
	createTableElement,
	createTextElement,
} from '../../../../src/core/pptx/core/builders/sdk/index';
import type {
	ChartInput,
	GroupOptions,
	ImageOptions,
	ShapeOptions,
	TableInput,
	TextOptions,
	TextSegmentInput,
	TextStyleInput,
	TransitionInput,
} from '../../../../src/core/pptx/core/builders/sdk/index';
import type { PptxHandler } from '../../../../src/core/pptx/core/PptxHandler';
import type { PptxElementAnimation } from '../../../../src/core/pptx/core/types/animation';
import type { PptxChartType } from '../../../../src/core/pptx/core/types/chart';
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import type { PptxSlide } from '../../../../src/core/pptx/core/types/presentation';
import type { PptxSmartArtData } from '../../../../src/core/pptx/core/types/smart-art';
import type { TextSegment } from '../../../../src/core/pptx/core/types/text';
import type { Box } from './geometry';
import type { Fx } from './timing';

export interface Run {
	text: string;
	style?: Partial<TextStyleInput>;
}
/** One paragraph: plain text, or runs with their own styles. */
export type Paragraph = string | Run[];

/** Text box options, plus the space before and after each paragraph (px). */
export type TextBoxOptions = TextOptions & { spaceBefore?: number; spaceAfter?: number };

type TextPlaceholder = 'title' | 'ctrTitle' | 'subTitle' | 'body' | 'sldNum';
type Animation = Omit<PptxElementAnimation, 'elementId'>;

const BREAK = '\n';

/** Paragraphs as SDK segments, with a marker segment between them. */
function segments(paragraphs: Paragraph[]): TextSegmentInput[] {
	const out: TextSegmentInput[] = [];
	paragraphs.forEach((paragraph, index) => {
		const runs = typeof paragraph === 'string' ? [{ text: paragraph }] : paragraph;
		if (index > 0) out.push({ text: BREAK, ...(runs[0]?.style ? { style: runs[0].style } : {}) });
		out.push(...runs);
	});
	return out;
}

/** Turns the marker segments into the model's paragraph breaks. */
function markBreaks(list: TextSegment[] | undefined): void {
	for (const segment of list ?? []) if (segment.text === BREAK) segment.isParagraphBreak = true;
}

export function dataUrl(png: Uint8Array): string {
	return `data:image/png;base64,${Buffer.from(png).toString('base64')}`;
}

export class Slide {
	readonly builder: SlideBuilder;
	/** Effects for the hand-written paragraph-build timing (see `timing.ts`). */
	readonly effects: Fx[] = [];
	private count = 0;

	constructor(
		readonly number: number,
		layoutPath: string,
		layoutName: string,
		handler: PptxHandler,
	) {
		this.builder = new SlideBuilder(number, layoutPath, layoutName, handler);
	}

	/** Adds an element under a name that is unique on the slide. */
	add<T extends PptxElement>(name: string, element: T): T {
		this.count += 1;
		element.id = `s${this.number}e${this.count}`;
		element.name = name;
		this.builder.addElement(element);
		return element;
	}

	/** Fills a layout placeholder; frame and text style come from the layout. */
	placeholder(type: TextPlaceholder, text: string | Paragraph[], idx?: number): PptxElement {
		const content = typeof text === 'string' ? text : segments(text);
		this.builder.addPlaceholder(type, content, idx === undefined ? {} : { idx });
		const element = this.builder.getLastElement()!;
		this.count += 1;
		element.id = `s${this.number}e${this.count}`;
		element.name = `${type[0]!.toUpperCase()}${type.slice(1)} ${this.count}`;
		if ('textSegments' in element) {
			markBreaks(element.textSegments);
			// A slide number is a field, so PowerPoint keeps it current when slides move.
			if (type === 'sldNum')
				for (const segment of element.textSegments ?? []) {
					segment.fieldType = 'slidenum';
					segment.fieldGuid = '{B1E0F6E2-3A51-4C0B-9D2C-5A1E7F3C9A10}';
				}
		}
		return element;
	}

	text(name: string, content: string | Paragraph[], options: TextBoxOptions): PptxElement {
		const { spaceBefore, spaceAfter, ...rest } = options;
		const element = createTextElement(
			typeof content === 'string' ? content : segments(content),
			rest,
		);
		markBreaks(element.textSegments);
		// Paragraph spacing is a property of the whole box in the SDK's text model.
		for (const style of [element.textStyle, ...(element.textSegments ?? []).map((s) => s.style)]) {
			if (style && spaceBefore !== undefined) style.paragraphSpacingBefore = spaceBefore;
			if (style && spaceAfter !== undefined) style.paragraphSpacingAfter = spaceAfter;
		}
		return this.add(name, element);
	}

	shape(name: string, geometry: string, options: ShapeOptions): PptxElement {
		const { text, ...rest } = options;
		const content = Array.isArray(text) ? text : text === undefined ? undefined : [{ text }];
		const element = createShapeElement(geometry, content ? { ...rest, text: content } : rest, {
			fill: { type: 'none' },
		});
		markBreaks(element.textSegments);
		return this.add(name, element);
	}

	/** A custom-geometry shape from SVG path data. */
	freeform(name: string, path: string, options: ShapeOptions): PptxElement {
		return this.add(name, createFreeformElement(path, options));
	}

	picture(name: string, png: Uint8Array, options: ImageOptions): PptxElement {
		return this.add(name, createImageElement(dataUrl(png), options));
	}

	chart(name: string, type: PptxChartType, input: ChartInput, box: Box): PptxElement {
		return this.add(name, createChartElement(type, input, box));
	}

	table(name: string, input: TableInput, box: Box): PptxElement {
		return this.add(name, createTableElement(input, box));
	}

	/** A SmartArt graphic; the core fabricates its five diagram parts on save. */
	smartArt(name: string, data: PptxSmartArtData, box: Box): PptxElement {
		return this.add(name, { type: 'smartArt', id: '', ...box, smartArtData: data });
	}

	group(name: string, children: PptxElement[], options?: GroupOptions): PptxElement {
		return this.add(name, createGroupElement(children, options));
	}

	/** An animation written by the core's animation writer. */
	animate(target: PptxElement, animation: Animation): this {
		const slide = this.builder.build();
		slide.animations ??= [];
		slide.animations.push({ elementId: target.id, ...animation });
		return this;
	}

	transition(input: TransitionInput): this {
		this.builder.setTransition(input);
		return this;
	}

	notes(text: string): this {
		this.builder.setNotes(text);
		return this;
	}

	get model(): PptxSlide {
		return this.builder.build();
	}
}
