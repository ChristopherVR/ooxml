import type { PptxMasterTextStyles, PptxSlide } from '../../types';
import type { PptxHandlerSaveOptions } from '../types';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSaveHandoutInfrastructure';

/** Serialises the deck being saved as a Transitional `.pptx` (see `saveAsLegacyPpt`). */
export type LegacyPptxSerializer = () => Promise<Uint8Array>;

/** True when `path` names a WAV file by extension (the only audio format `media-writer.ts` embeds). */
function isWavPath(path: string): boolean {
	return /\.wav$/iu.test(path);
}

/**
 * The `'ppt'` output-format branch of the save pipeline.
 *
 * A legacy binary `.ppt` is not an OOXML ZIP package at all, so this bypasses
 * every other save mixin entirely (no XML builder, no OOXML relationship
 * bookkeeping) rather than reusing the OOXML pipeline and converting its
 * output: it builds the `write-ppt.ts` intermediate model directly from the
 * live `PptxSlide[]` and hands it to the binary writer. `this.zip` (the
 * loaded archive, still populated even on this branch: see
 * `PptxHandlerRuntimeLoadSession.ts`) is read from, never written to, purely
 * to resolve `media`-`audio` elements loaded from a real `.pptx` (see
 * `resolveAudioMediaBytes`'s doc).
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * A `media`-`audio` element loaded from a real `.pptx` carries only
	 * `mediaPath` (a lazy reference into `this.zip`), never `mediaData`
	 * (unlike `imageData`, which the loader eagerly base64-encodes): see
	 * `PptxSlideLoaderService.ts`. Resolve every such path's WAV bytes here,
	 * before conversion, so `element-to-write-model.ts`'s `ConvertContext.
	 * resolvedMedia` can embed them without mutating the live element tree
	 * (the OOXML save pipeline's `PptxHandlerRuntimeSaveElementEmbedding.ts`
	 * reads from the exact same `this.zip` path for the equivalent case, but
	 * mutates the live element in place, which this branch avoids since
	 * nothing else about the live model is touched by a `.ppt` save).
	 */
	private async resolveAudioMediaBytes(slides: PptxSlide[]): Promise<Map<string, Uint8Array>> {
		const paths = new Set<string>();
		const walk = (elements: PptxSlide['elements']): void => {
			for (const el of elements) {
				if (
					el.type === 'media' &&
					el.mediaType === 'audio' &&
					!el.mediaData &&
					el.mediaPath &&
					isWavPath(el.mediaPath)
				) {
					paths.add(el.mediaPath);
				} else if (el.type === 'group') {
					walk(el.children);
				}
			}
		};
		for (const slide of slides) {
			walk(slide.elements);
		}
		const resolved = new Map<string, Uint8Array>();
		await Promise.all(
			[...paths].map(async (path) => {
				const bytes = await this.zip.file(path)?.async('uint8array');
				if (bytes) {
					resolved.set(path, bytes);
				}
			}),
		);
		return resolved;
	}

	/**
	 * The text styles of the master at `masterPath` (the first slide's master
	 * when omitted), preferring an edited copy passed in `options.slideMasters`
	 * over the styles parsed at load.
	 */
	private legacyMasterTextStyles(
		slides: PptxSlide[],
		options: PptxHandlerSaveOptions | undefined,
		masterPath?: string,
	): PptxMasterTextStyles | undefined {
		const layoutPath = slides[0]?.layoutPath;
		const path =
			(masterPath && this.masterTxStylesCache.has(masterPath) ? masterPath : undefined) ??
			(layoutPath ? this.resolveMasterPathForLayout(layoutPath) : undefined) ??
			this.masterTxStylesCache.keys().next().value;
		if (!path) {
			return undefined;
		}
		const edited = options?.slideMasters?.find((master) => master.path === path)?.txStyles;
		return edited ?? this.masterTxStylesCache.get(path);
	}

	/**
	 * Serialise `slides` to a legacy binary PowerPoint 97-2003 (`.ppt`) file.
	 *
	 * @param slides - The (possibly mutated) live slide array.
	 * @param options - `pptPassword` and `rasterizeSvg` (the `.ppt`-only
	 *   options), `customShows` and `slideMasters`; every OOXML-specific
	 *   option (conformance, embedded fonts, table styles, ...) has no binary
	 *   `.ppt` counterpart and is ignored.
	 * @param saveAsPptx - Serialises the same slides as a Transitional `.pptx`
	 *   (the save pipeline's own OOXML path), the source of the `metroBlob`
	 *   packages and of the master's placeholders and theme round-trip atoms;
	 *   omit to write plain fallbacks and a default master only.
	 */
	protected async saveAsLegacyPpt(
		slides: PptxSlide[],
		options: PptxHandlerSaveOptions | undefined,
		saveAsPptx?: LegacyPptxSerializer,
	): Promise<Uint8Array> {
		const {
			buildDeckMasterRoundTripsFromPptx,
			buildMetroBlobs,
			buildPptFile,
			convertDeckToWriteModel,
			deckNeedsMetroBlobs,
			convertMasterTextStyles,
			resolvePictureSources,
		} = await import('../../ppt/writer');
		// One Transitional .pptx save feeds both the metroBlob packages and each
		// main master's placeholders, theme and text-style round-trip atoms.
		const pptxBytes = saveAsPptx ? await saveAsPptx() : undefined;
		const metroBlobs =
			pptxBytes && deckNeedsMetroBlobs(slides)
				? await buildMetroBlobs(pptxBytes, slides)
				: undefined;
		const deckMasters = pptxBytes ? await buildDeckMasterRoundTripsFromPptx(pptxBytes) : undefined;
		this.compatibilityService.resetWarnings();
		const resolvedMedia = await this.resolveAudioMediaBytes(slides);
		const resolvedPictures = await resolvePictureSources(
			slides,
			async (path) => this.zip.file(path)?.async('uint8array'),
			{ rasterizeSvg: options?.rasterizeSvg },
		);
		const deck = convertDeckToWriteModel(
			slides,
			this.rawSlideWidthEmu || 9144000,
			this.rawSlideHeightEmu || 6858000,
			(warning) => this.compatibilityService.reportWarning(warning),
			options?.customShows,
			resolvedMedia,
			metroBlobs,
			resolvedPictures,
		);
		if (deckMasters && deckMasters.masterPaths.length > 1) {
			// One .ppt main master per source master a slide uses, each slide
			// pointing at its own (see `buildDeckMasterRoundTripsFromPptx`).
			deck.masters = deckMasters.masters.map((roundTrip, i) => {
				const styles = convertMasterTextStyles(
					this.legacyMasterTextStyles(slides, options, deckMasters.masterPaths[i]),
					roundTrip.themeFonts,
				);
				return styles ? { roundTrip, styles } : { roundTrip };
			});
			deck.slides.forEach((slide, i) => {
				const masterIndex = deckMasters.slideMasterIndex[i];
				if (masterIndex !== undefined) slide.masterIndex = masterIndex;
			});
		} else {
			const master = deckMasters?.masters[0];
			const masterStyles = convertMasterTextStyles(
				this.legacyMasterTextStyles(slides, options, deckMasters?.masterPaths[0]),
				master?.themeFonts,
			);
			if (masterStyles) deck.masterStyles = masterStyles;
			if (master) deck.master = master;
		}
		return buildPptFile(deck, {
			...(options?.pptPassword !== undefined && { password: options?.pptPassword }),
		});
	}
}
