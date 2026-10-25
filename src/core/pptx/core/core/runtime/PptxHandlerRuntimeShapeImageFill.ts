import { PptxElement, ShapeStyle, XmlObject, TextStyle } from '../../types';
import { resolveGroupFillImagePure, type GroupFillImageHost } from './group-parsing-helpers';
import { xmlAttr, xmlChild, xmlPath } from '../../utils/xml-access';
import { parseEmuInt } from './group-shape-geometry';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeGeometryParsing';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Parse a shape that has an image fill (a:blipFill inside spPr)
	 * This handles shapes like rectangles filled with images (e.g., wood texture backgrounds)
	 */
	protected async parseShapeWithImageFill(
		shape: XmlObject,
		id: string,
		slidePath: string,
	): Promise<PptxElement | null> {
		try {
			const spPr = xmlChild(shape, 'p:spPr');
			const placeholderInfo = this.extractPlaceholderInfo(xmlPath(shape, 'p:nvSpPr', 'p:nvPr'));
			const inheritedPlaceholder = placeholderInfo
				? this.findPlaceholderContext(slidePath, placeholderInfo)
				: undefined;
			const inheritedSpPr = (inheritedPlaceholder?.shape?.['p:spPr'] ||
				inheritedPlaceholder?.picture?.['p:spPr']) as XmlObject | undefined;
			const effectiveSpPr = this.mergeXmlObjects(inheritedSpPr, spPr);
			const xfrm = (effectiveSpPr?.['a:xfrm'] || spPr?.['a:xfrm'] || inheritedSpPr?.['a:xfrm']) as
				| XmlObject
				| undefined;
			if (!xfrm) {
				return null;
			}

			const off = xmlChild(xfrm, 'a:off');
			const ext = xmlChild(xfrm, 'a:ext');
			if (!off || !ext) {
				return null;
			}

			// Capture the exact source EMU alongside the rounded pixel value so
			// `resolveXfrmEmu` can re-emit a byte-identical `a:off`/`a:ext` for this
			// shape on save when it was never moved or resized (see
			// `xfrm-emu-resolution.ts`). Without these, EVERY shape-with-image-fill
			// re-quantized from pixels on every save, drifting up to +/-4762 EMU
			// (half a pixel) with no edit involved: `image-fill-rot-with-shape.pptx`
			// round-trips 508000 -> 504825 on an unrelated sibling shape moving.
			const xEmu = parseEmuInt(xmlAttr(off, 'x'));
			const yEmu = parseEmuInt(xmlAttr(off, 'y'));
			const widthEmu = parseEmuInt(xmlAttr(ext, 'cx'));
			const heightEmu = parseEmuInt(xmlAttr(ext, 'cy'));
			const x = Math.round(xEmu / PptxHandlerRuntime.EMU_PER_PX);
			const y = Math.round(yEmu / PptxHandlerRuntime.EMU_PER_PX);
			const width = Math.round(widthEmu / PptxHandlerRuntime.EMU_PER_PX);
			const height = Math.round(heightEmu / PptxHandlerRuntime.EMU_PER_PX);

			// Get rotation if present
			const rotation = xfrm['@_rot'] ? parseInt(xfrm['@_rot']) / 60000 : undefined;
			const skewX = xfrm['@_skewX'] ? parseInt(String(xfrm['@_skewX']), 10) / 60000 : undefined;
			const skewY = xfrm['@_skewY'] ? parseInt(String(xfrm['@_skewY']), 10) / 60000 : undefined;
			const { flipHorizontal, flipVertical } = this.readFlipState(xfrm);

			const prstGeom = xmlAttr(xmlChild(effectiveSpPr, 'a:prstGeom'), 'prst');
			const shapeAdjustments = this.parseGeometryAdjustments(
				effectiveSpPr?.['a:prstGeom'] as XmlObject | undefined,
			);
			let shapeType = prstGeom || 'rect';
			let pathData: string | undefined;
			let pathWidth: number | undefined;
			let pathHeight: number | undefined;

			const custGeom = effectiveSpPr?.['a:custGeom'] as XmlObject | undefined;
			if (custGeom) {
				const customPath = this.parseCustomGeometry(custGeom, width, height);
				if (customPath) {
					shapeType = 'custom';
					pathData = customPath.pathData;
					pathWidth = customPath.pathWidth;
					pathHeight = customPath.pathHeight;
				}
			}

			// Parse adjustment handles from geometry definition
			const geomNode = custGeom ?? (effectiveSpPr?.['a:prstGeom'] as XmlObject | undefined);
			const adjustmentHandles = this.parseAdjustmentHandles(
				geomNode,
				width,
				height,
				shapeAdjustments,
			);

			// Get image relationship ID from spPr's blipFill (not p:blipFill)
			const blipFill = (effectiveSpPr?.['a:blipFill'] || spPr?.['a:blipFill']) as
				| XmlObject
				| undefined;
			const blip = xmlChild(blipFill, 'a:blip');
			const rEmbed = xmlAttr(blip, 'r:embed');
			const rLink = xmlAttr(blip, 'r:link');
			const crop = this.readImageCropFromBlipFill(blipFill);
			// A shape image fill carries the same blip effects as a regular p:pic.
			// Dropping these made authored transparency such as
			// <a:alphaModFix amt="15000"/> render at full opacity.
			const imageEffects = this.extractImageEffects(blip);

			// Image tiling properties from a:tile
			const tileNode = blipFill?.['a:tile'] as XmlObject | undefined;
			const tileProps: Record<string, unknown> = {};
			if (tileNode) {
				const txRaw = Number.parseInt(String(tileNode['@_tx'] || ''), 10);
				if (Number.isFinite(txRaw)) {
					tileProps.tileOffsetX = txRaw / PptxHandlerRuntime.EMU_PER_PX;
				}
				const tyRaw = Number.parseInt(String(tileNode['@_ty'] || ''), 10);
				if (Number.isFinite(tyRaw)) {
					tileProps.tileOffsetY = tyRaw / PptxHandlerRuntime.EMU_PER_PX;
				}
				const sxRaw = Number.parseInt(String(tileNode['@_sx'] || ''), 10);
				if (Number.isFinite(sxRaw)) {
					tileProps.tileScaleX = sxRaw / 100000;
				}
				const syRaw = Number.parseInt(String(tileNode['@_sy'] || ''), 10);
				if (Number.isFinite(syRaw)) {
					tileProps.tileScaleY = syRaw / 100000;
				}
				const flipStr = String(tileNode['@_flip'] || '').trim();
				if (flipStr === 'x' || flipStr === 'y' || flipStr === 'xy' || flipStr === 'none') {
					tileProps.tileFlip = flipStr;
				}
				const algnStr = String(tileNode['@_algn'] || '').trim();
				if (algnStr.length > 0) {
					tileProps.tileAlignment = algnStr;
				}
			}

			// Print-resolution hint (`a:blipFill/@dpi`): round-trip only, no
			// on-screen rendering effect. See `PptxImageProperties.dpi`.
			const dpiRaw = Number.parseInt(String(blipFill?.['@_dpi'] || ''), 10);
			if (Number.isFinite(dpiRaw) && dpiRaw > 0) {
				tileProps.dpi = dpiRaw;
			}

			this.compatibilityService.inspectPictureCompatibility(blipFill, blip, slidePath, id);
			this.compatibilityService.inspectShapeCompatibility(
				effectiveSpPr,
				shape['p:txBody'] as XmlObject | undefined,
				slidePath,
				id,
			);

			let imageData: string | undefined;
			let imagePath: string | undefined;

			const relId = rEmbed || rLink;
			if (relId) {
				const slideRels = this.slideRelsMap.get(slidePath);
				const target = slideRels?.get(relId);

				if (target) {
					if (target.startsWith('http://') || target.startsWith('https://')) {
						// A linked (TargetMode="External") picture fill is only referenced when
						// the host allows remote content, like every other blip.
						if (this.allowExternalImages === true) {
							imagePath = target;
							imageData = target;
						}
					} else if (target.startsWith('data:')) {
						imagePath = target;
						imageData = target;
					} else {
						imagePath = this.resolveImagePath(slidePath, target);
						if (this.eagerDecodeImages && imagePath) {
							imageData = await this.getImageData(imagePath);
						}
					}
				}
			}
			const styleNode = (shape['p:style'] ||
				inheritedPlaceholder?.shape?.['p:style'] ||
				inheritedPlaceholder?.picture?.['p:style']) as XmlObject | undefined;

			// Parse hyperlink / action for the shape-with-image-fill element
			const sifCNvPr = xmlPath(shape, 'p:nvSpPr', 'p:cNvPr');
			const sifElementName = xmlAttr(sifCNvPr, 'name')?.trim() || undefined;
			const sifSlideRels = this.slideRelsMap.get(slidePath);
			const { actionClick: sifActionClick, actionHover: sifActionHover } = this.parseElementActions(
				sifCNvPr,
				sifSlideRels,
				this.orderedSlidePaths,
			);

			return {
				id,
				type: 'picture',
				x,
				y,
				width,
				height,
				xEmu,
				yEmu,
				widthEmu,
				heightEmu,
				imageData,
				imagePath,
				imageEffects: imageEffects || undefined,
				...crop,
				...tileProps,
				shapeType,
				shapeAdjustments,
				adjustmentHandles,
				pathData,
				pathWidth,
				pathHeight,
				shapeStyle: this.extractShapeStyle(effectiveSpPr, styleNode),
				rotation,
				skewX,
				skewY,
				flipHorizontal,
				flipVertical,
				rawXml: shape,
				name: sifElementName,
				actionClick: sifActionClick,
				actionHover: sifActionHover,
			};
		} catch (e) {
			console.warn(`[pptx] Skipping shape-with-image-fill (${id}):`, e);
			return null;
		}
	}

	/**
	 * A text-bearing shape whose picture fill is inherited from its layout or
	 * master placeholder (`a:blipFill` absent from its own `spPr`) gets
	 * `fillMode: 'image'` from the merged properties but no image: the blip's
	 * relationship id belongs to the part that authored it, not the slide.
	 * Resolve it against the layout's, then the master's, relationships.
	 */
	protected async resolveInheritedFillImage(
		element: PptxElement | null,
		shape: XmlObject,
		slidePath: string,
	): Promise<void> {
		const style = (element as { shapeStyle?: ShapeStyle } | null)?.shapeStyle;
		if (!element || !style || style.fillMode !== 'image' || style.fillImageUrl) {
			return;
		}
		const placeholderInfo = this.extractPlaceholderInfo(xmlPath(shape, 'p:nvSpPr', 'p:nvPr'));
		const inherited = placeholderInfo
			? this.findPlaceholderContext(slidePath, placeholderInfo)
			: undefined;
		const effectiveSpPr = this.mergeXmlObjects(
			inherited?.shape?.['p:spPr'] as XmlObject | undefined,
			xmlChild(shape, 'p:spPr'),
		);
		const blipFill = effectiveSpPr?.['a:blipFill'] as XmlObject | undefined;
		const blip = blipFill?.['a:blip'] as XmlObject | undefined;
		const relId = String(blip?.['@_r:embed'] ?? blip?.['@_r:link'] ?? '');
		if (!relId) {
			return;
		}
		const layoutPath = this.resolveLayoutPathForSlide(slidePath);
		const masterPath = layoutPath ? this.resolveMasterPathForLayout(layoutPath) : undefined;
		for (const partPath of [layoutPath, masterPath]) {
			if (!partPath || !this.slideRelsMap.get(partPath)?.has(relId)) {
				continue;
			}
			const fill = await resolveGroupFillImagePure(
				this as unknown as GroupFillImageHost,
				blipFill,
				partPath,
			);
			if (fill) {
				// `fillImageUrl` has no lazy-load patch target (unlike a picture's
				// `imagePath`/`imageData` pair), so an unresolved archive path would
				// 404 in the browser: decode it now whatever `eagerDecodeImages` says.
				const isUrl = /^(?:data:|blob:|https?:)/u.test(fill.fillImageUrl);
				const decoded = isUrl ? undefined : await this.getImageData(fill.fillImageUrl);
				style.fillImageUrl = decoded ?? fill.fillImageUrl;
				style.fillImageMode = fill.fillImageMode;
				// Painted fill makes it a styled shape, as a solid inherited fill does.
				if (element.type === 'text') {
					(element as { type: PptxElement['type'] }).type = 'shape';
				}
			}
			return;
		}
	}

	protected textVerticalAlignFromDrawingValue(value: unknown): TextStyle['vAlign'] | undefined {
		const normalized = String(value ?? '')
			.trim()
			.toLowerCase();
		if (normalized.length === 0) {
			return undefined;
		}
		if (normalized === 't' || normalized === 'top') {
			return 'top';
		}
		if (normalized === 'ctr' || normalized === 'center') {
			return 'middle';
		}
		if (normalized === 'b' || normalized === 'bottom') {
			return 'bottom';
		}
		// D1-G5 (D2 audit numbering: text D2-G5): `dist`/`just` are distinct
		// from true centering (ST_TextAnchoringType, ECMA-376 §20.1.10.2) - keep
		// them as their own values so save can round-trip the original
		// attribute instead of silently rewriting it to `ctr`.
		if (normalized === 'dist') {
			return 'distributed';
		}
		if (normalized === 'just') {
			return 'justified';
		}
		return undefined;
	}

	protected textDirectionFromDrawingValue(value: unknown): TextStyle['textDirection'] | undefined {
		const normalized = String(value ?? '')
			.trim()
			.toLowerCase();
		if (normalized.length === 0 || normalized === 'horz') {
			return undefined;
		}
		if (normalized === 'vert') {
			return 'vertical';
		}
		if (normalized === 'vert270') {
			return 'vertical270';
		}
		if (normalized === 'eavert') {
			return 'eaVert';
		}
		if (normalized === 'wordartvert') {
			return 'wordArtVert';
		}
		if (normalized === 'wordartvertrtl') {
			return 'wordArtVertRtl';
		}
		if (normalized === 'mongolianvert') {
			return 'mongolianVert';
		}
		return undefined;
	}
}
