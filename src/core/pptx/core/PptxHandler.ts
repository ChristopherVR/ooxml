import {
	applyImportedPptxData,
	decodePptxJsonText,
	deserializePptxFromJson,
} from '../converter/json';
import {
	PresentationBuilder,
	buildBlankPresentationArchive,
} from './builders/sdk/PresentationBuilder';
import type { PresentationBuilderResult } from './builders/sdk/PresentationBuilder';
import type { PresentationOptions } from './builders/sdk/types';
import type { PptxHandlerLoadOptions } from './core';
import {
	assertPptxInputWithinLimits,
	assertPptxSlideCountWithinLimits,
	countPptxShapesWithinLimits,
	resolvePptxLoadLimits,
} from './core/load-limits';
import { PptxHandlerCore } from './PptxHandlerCore';
import type { PptxData } from './types';

/**
 * Public facade for the PPTX editor handler.
 *
 * The implementation lives in `PptxHandlerCore` so this surface can stay small,
 * stable, and easy to replace with alternate implementations in the future.
 */
export class PptxHandler extends PptxHandlerCore {
	/**
	 * Create a new blank PPTX presentation from scratch.
	 *
	 * This is a convenience static method that delegates to
	 * {@link PresentationBuilder.create}. The returned handler is fully
	 * initialized and ready for editing, adding slides, and saving.
	 *
	 * @param options - Optional slide dimensions, theme, and metadata.
	 * @returns Handler, parsed data, and a slide builder factory.
	 *
	 * @example
	 * ```ts
	 * const { handler, data, createSlide } = await PptxHandler.createBlank({
	 *   title: "My Deck",
	 *   theme: { colors: { accent1: "#FF6B6B" } },
	 * });
	 *
	 * data.slides.push(
	 *   createSlide("Blank")
	 *     .addText("Hello", { fontSize: 36 })
	 *     .build()
	 * );
	 *
	 * const bytes = await handler.save(data.slides);
	 * ```
	 */
	static async createBlank(options?: PresentationOptions): Promise<PresentationBuilderResult> {
		return PresentationBuilder.create(options);
	}

	/**
	 * Create a new PPTX presentation from scratch.
	 *
	 * Alias for {@link createBlank}. Generates a valid minimal OpenXML
	 * package and returns a fully initialized handler ready for editing,
	 * adding slides, and saving.
	 *
	 * @param options - Optional slide dimensions, theme, metadata,
	 *   and initial slide count.
	 * @returns Handler, parsed data, and a slide builder factory.
	 *
	 * @example
	 * ```ts
	 * const { handler, data, createSlide } = await PptxHandler.create({
	 *   title: "Q4 Report",
	 *   initialSlideCount: 3,
	 *   theme: { colors: { accent1: "#FF6B6B" } },
	 * });
	 *
	 * // The presentation already has 3 blank slides
	 * console.log(data.slides.length); // => 3
	 *
	 * // Add more slides with content
	 * data.slides.push(
	 *   createSlide("Blank")
	 *     .addText("Hello", { fontSize: 36 })
	 *     .build()
	 * );
	 *
	 * const bytes = await handler.save(data.slides);
	 * ```
	 */
	static async create(options?: PresentationOptions): Promise<PresentationBuilderResult> {
		return PresentationBuilder.create(options);
	}

	/**
	 * Parse a presentation from an `ArrayBuffer`, accepting both `.pptx`
	 * archives and portable `pptx-viewer-json` documents.
	 *
	 * The buffer is sniffed first: JSON documents (leading `{` plus the
	 * `"pptx-viewer-json"` format marker) are routed to {@link loadFromJson};
	 * everything else goes through the regular ZIP/OLE2 pipeline.
	 */
	public override async load(
		data: ArrayBuffer,
		options: PptxHandlerLoadOptions = {},
	): Promise<PptxData> {
		// Before the JSON sniff too, so `maxInputBytes` also bounds a JSON input.
		assertPptxInputWithinLimits(data.byteLength, resolvePptxLoadLimits(options));
		const jsonText = decodePptxJsonText(data);
		if (jsonText !== null) {
			return this.loadFromJson(jsonText, options);
		}
		return super.load(data, options);
	}

	/**
	 * Load a presentation from `pptx-viewer-json` text.
	 *
	 * A minimal blank package is generated and loaded first so that the
	 * handler keeps a valid in-memory archive (editing and {@link save} keep
	 * working), then the imported model is overlaid on top: imported
	 * presentation fields win, and the slide array is replaced wholesale.
	 */
	public async loadFromJson(text: string, options: PptxHandlerLoadOptions = {}): Promise<PptxData> {
		const limits = resolvePptxLoadLimits(options);
		const imported = deserializePptxFromJson(text);
		// The imported slides never pass through the package loader, so the
		// slide and shape limits are applied to them here, before they are used.
		assertPptxSlideCountWithinLimits(imported.slides.length, limits);
		let shapes = 0;
		for (const slide of imported.slides) {
			shapes = countPptxShapesWithinLimits(shapes, slide.elements ?? [], limits);
		}
		const blankBytes = await buildBlankPresentationArchive({
			width: imported.widthEmu,
			height: imported.heightEmu,
			// Saving writes presentation.xml from the loaded archive, not from
			// the overlaid model, so the archive must carry the value itself.
			firstSlideNumber: imported.firstSlideNumber,
		});
		// The limits were applied to the imported model above; the generated
		// blank package is not the host's input, so it loads with the defaults.
		const base = await super.load(blankBytes, { ...options, limits: undefined });
		return applyImportedPptxData(base, imported);
	}
}
