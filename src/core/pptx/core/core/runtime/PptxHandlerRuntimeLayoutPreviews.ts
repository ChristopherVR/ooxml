import { EMU_PER_PX } from '../../constants';
import type { PptxLayoutPreview, XmlObject } from '../../types';
import { xmlChild, xmlPath } from '../../utils/xml-access';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeLayoutElements';

/**
 * Layout thumbnail sourcing for the New Slide / Layout galleries.
 *
 * Previews are produced on request rather than during load. Materialising
 * every layout's artwork means parsing each `p:sldLayout` part and decoding
 * the images it references, which is measurable on decks carrying several
 * masters, and the majority of sessions never open the gallery at all. The
 * results are memoised, so opening the gallery a second time is free.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	private layoutPreviewCache: Map<string, PptxLayoutPreview> = new Map();

	/**
	 * Build (or return the memoised) preview for one layout.
	 *
	 * @param layoutPath - Archive path of the `p:sldLayout` part.
	 * @returns The preview, or `null` when the archive holds no such layout.
	 */
	public async getLayoutPreview(layoutPath: string): Promise<PptxLayoutPreview | null> {
		const cached = this.layoutPreviewCache.get(layoutPath);
		if (cached) {
			return cached;
		}

		const layoutXml = this.layoutXmlMap.get(layoutPath);
		if (!layoutXml) {
			return null;
		}

		// Thumbnails are worthless without their pictures, and a layout parsed
		// earlier on behalf of a slide may have been cached with its images left
		// as archive references. Drop that entry so this pass re-reads it with
		// decoding on; the decoded result is a superset, so slide rendering
		// keeps working from the same cache afterwards.
		const previousEagerDecode = this.eagerDecodeImages;
		if (!previousEagerDecode) {
			this.layoutCache.delete(layoutPath);
			this.eagerDecodeImages = true;
		}

		// Scheme colours, the background's included, resolve through this
		// layout's own master, not whichever master a slide parse left active.
		const masterPath = this.resolveMasterPathForLayout(layoutPath);
		let parsed;
		try {
			parsed = await this.withMasterThemeState(masterPath, async () => ({
				elements: await this.getLayoutElementsByPath(layoutPath),
				background: this.resolveLayoutPreviewBackground(layoutXml, masterPath),
			}));
		} finally {
			this.eagerDecodeImages = previousEagerDecode;
		}

		const cSld = xmlPath(layoutXml, 'p:sldLayout', 'p:cSld');

		const preview: PptxLayoutPreview = {
			path: layoutPath,
			width: this.rawSlideWidthEmu / EMU_PER_PX,
			height: this.rawSlideHeightEmu / EMU_PER_PX,
			...parsed.background,
			elements: parsed.elements,
			placeholders: this.extractPlaceholderList(xmlChild(cSld, 'p:spTree')),
		};

		this.layoutPreviewCache.set(layoutPath, preview);
		return preview;
	}

	/**
	 * Build previews for every layout the presentation defines.
	 *
	 * The default order matches `getLayoutOptions`, since both walk the layout
	 * XML map, which the master parser fills in `p:sldLayoutIdLst` order.
	 *
	 * @param layoutPaths - Restricts the result to these layouts; defaults to
	 *   all of them, which is what an unscoped gallery wants.
	 */
	public async getLayoutPreviews(layoutPaths?: readonly string[]): Promise<PptxLayoutPreview[]> {
		const paths = layoutPaths ?? [...this.layoutXmlMap.keys()];
		const previews: PptxLayoutPreview[] = [];
		// Sequential on purpose: each parse mutates shared colour-map state on
		// the runtime, so overlapping them would interleave those mutations.
		for (const path of paths) {
			const preview = await this.getLayoutPreview(path);
			if (preview) {
				previews.push(preview);
			}
		}
		return previews;
	}

	/** Drop memoised previews, e.g. after a layout is edited in template mode. */
	protected invalidateLayoutPreviews(layoutPath?: string): void {
		if (layoutPath) {
			this.layoutPreviewCache.delete(layoutPath);
		} else {
			this.layoutPreviewCache.clear();
		}
	}

	/**
	 * Resolve the background a layout thumbnail should paint, falling back to
	 * the layout's master when the layout itself declares none. Runs with the
	 * master active; the layout's own colour map override applies to an
	 * inherited background too, as a slide's does.
	 */
	private resolveLayoutPreviewBackground(
		layoutXml: XmlObject,
		masterPath: string | undefined,
	): Pick<PptxLayoutPreview, 'backgroundColor' | 'backgroundImage'> {
		const clrMapOverride = this.parseLayoutClrMapOverride(layoutXml);
		const backgroundColor =
			this.extractTemplateBackgroundColor(layoutXml, 'p:sldLayout', clrMapOverride) ??
			this.extractTemplateBackgroundColor(
				masterPath ? this.masterXmlMap.get(masterPath) : undefined,
				'p:sldMaster',
				clrMapOverride,
			);
		return backgroundColor ? { backgroundColor } : {};
	}
}
