import type { Model3DPptxElement, XmlObject } from '../../types';
import { isExternalTarget } from '../builders/PptxSlideRelationshipRegistry';
import { resolveXfrmEmu } from '../../utils/xfrm-emu-resolution';
import { ensureXmlChildOrCreate } from '../../utils/xml-access';
import {
	defaultModel3DSceneXml,
	MODEL3D_REL,
	MODEL3D_URI,
	readModel3DGraphic,
} from '../../utils/model3d-writing';
import type { SaveSlideContext } from './PptxHandlerRuntimeSaveElementEmbedding';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSaveSummaryZoom';

const MC_NAMESPACE = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	private readonly newModel3DFallbackByXml = new Map<XmlObject, XmlObject>();

	/** Create or update an embedded GLB model, its poster, and package relationships. */
	protected createOrUpdateModel3DXml(
		el: Model3DPptxElement,
		shape: XmlObject | undefined,
		ctx: SaveSlideContext,
	): XmlObject | undefined {
		const native = readModel3DGraphic(shape);
		const existingModelRel = shape?.['p16:model3Drel'] as XmlObject | undefined;
		let modelRelationshipId =
			String(native?.['@_r:embed'] ?? existingModelRel?.['@_r:id'] ?? '').trim() || undefined;
		let modelPath = el.modelPath;
		if (typeof el.modelData === 'string') {
			const parsed = this.parseDataUrlToBytes(el.modelData);
			if (!parsed) {
				return shape;
			}
			if (!modelPath?.toLowerCase().endsWith('.glb')) {
				modelPath = ctx.saveSession.nextMediaPath('glb');
			}
			this.zip.file(modelPath, parsed.bytes);
		}
		if (!modelPath) {
			return shape;
		}
		modelRelationshipId ??= ctx.slideRelationshipRegistry.nextRelationshipId();
		ctx.slideRelationshipRegistry.upsertRelationship(
			modelRelationshipId,
			MODEL3D_REL,
			modelPath.replace(/^ppt\//u, '../'),
		);
		el.modelPath = modelPath;
		el.modelMimeType = 'model/gltf-binary';

		const poster = this.ensureModel3DPoster(el, shape, native, ctx);
		if (!shape) {
			shape = this.buildModel3DGraphicFrame(el, modelRelationshipId, poster.relationshipId);
			this.newModel3DFallbackByXml.set(
				shape,
				this.buildModel3DFallbackPicture(el, poster.relationshipId),
			);
			el.rawXml = shape;
			return shape;
		}

		if (native) {
			native['@_r:embed'] = modelRelationshipId;
			this.applyModel3DTransform(shape, el, 'p:xfrm');
			this.applyLocalModel3DSize(native, el);
			if (poster.relationshipId) {
				const raster = ensureXmlChildOrCreate(native, 'am3d:raster');
				ensureXmlChildOrCreate(raster, 'am3d:blip')['@_r:embed'] = poster.relationshipId;
				this.updateModel3DFallback(shape, el, poster.relationshipId);
			}
			return shape;
		}

		ensureXmlChildOrCreate(shape, 'p16:model3Drel')['@_r:id'] = modelRelationshipId;
		const spPr = ensureXmlChildOrCreate(shape, 'p16:spPr');
		this.applyModel3DTransform(spPr, el);
		if (poster.relationshipId) {
			const posterNode = ensureXmlChildOrCreate(shape, 'p16:posterImage');
			posterNode['@_r:embed'] = poster.relationshipId;
			this.updateModel3DFallback(shape, el, poster.relationshipId);
		}
		return shape;
	}

	/** Wrap newly authored graphic frames alongside their raster fallback. */
	protected wrapNewModel3DEnvelopes(spTree: XmlObject, models: readonly XmlObject[]): void {
		const envelopes: XmlObject[] = [];
		for (const model of models) {
			const fallback = this.newModel3DFallbackByXml.get(model);
			if (!fallback || this.alternateContentBlockByRawXml.has(model)) {
				continue;
			}
			envelopes.push({
				'@_xmlns:mc': MC_NAMESPACE,
				'mc:Choice': {
					'@_Requires': 'am3d',
					'@_xmlns:am3d': MODEL3D_URI,
					'p:graphicFrame': model,
				},
				'mc:Fallback': { 'p:pic': fallback },
			});
		}
		if (envelopes.length === 0) {
			return;
		}
		spTree['p:graphicFrame'] = models.filter((model) => !this.newModel3DFallbackByXml.has(model));
		const existing = this.ensureArray(spTree['mc:AlternateContent']) as XmlObject[];
		spTree['mc:AlternateContent'] = [...existing, ...envelopes];
	}

	private ensureModel3DPoster(
		el: Model3DPptxElement,
		shape: XmlObject | undefined,
		native: XmlObject | undefined,
		ctx: SaveSlideContext,
	): { relationshipId?: string | undefined } {
		const posterNode = shape?.['p16:posterImage'] as XmlObject | undefined;
		const raster = native?.['am3d:raster'] as XmlObject | undefined;
		const blip = raster?.['am3d:blip'] as XmlObject | undefined;
		let relationshipId =
			String(blip?.['@_r:embed'] ?? posterNode?.['@_r:embed'] ?? '').trim() || undefined;
		let imagePath = el.imagePath;
		const posterData =
			el.imageData ??
			(typeof el.posterImage === 'string' && el.posterImage.startsWith('data:')
				? el.posterImage
				: undefined);
		if (posterData) {
			const parsed = this.parseDataUrlToBytes(posterData);
			if (parsed) {
				imagePath ??= ctx.saveSession.nextMediaPath(parsed.extension);
				this.zip.file(imagePath, parsed.bytes);
			}
		} else if (!imagePath && el.posterImage?.startsWith('ppt/')) {
			imagePath = el.posterImage;
		}
		if (imagePath) {
			relationshipId ??= ctx.slideRelationshipRegistry.nextRelationshipId();
			ctx.slideRelationshipRegistry.upsertRelationship(
				relationshipId,
				ctx.slideImageRelationshipType,
				imagePath.replace(/^ppt\//u, '../'),
				isExternalTarget(imagePath) ? 'External' : undefined,
			);
			el.imagePath = imagePath;
		}
		return { relationshipId };
	}

	private buildModel3DGraphicFrame(
		el: Model3DPptxElement,
		relationshipId: string,
		previewId?: string,
	): XmlObject {
		const scene = this.parser.parse(`<scene>${defaultModel3DSceneXml(previewId)}</scene>`) as {
			scene: XmlObject;
		};
		const model: XmlObject = {
			'@_r:embed': relationshipId,
			'am3d:spPr': this.buildModel3DShapeProperties(el),
			...scene.scene,
		};
		this.applyLocalModel3DSize(model, el);
		const frame: XmlObject = {
			'p:nvGraphicFramePr': {
				'p:cNvPr': {
					'@_id': el.shapeId ?? '0',
					'@_name': el.name ?? el.id,
					...(el.altText ? { '@_descr': el.altText } : {}),
				},
				'p:cNvGraphicFramePr': {},
				'p:nvPr': {},
			},
			'p:xfrm': {},
			'a:graphic': { 'a:graphicData': { '@_uri': MODEL3D_URI, 'am3d:model3d': model } },
		};
		this.applyModel3DTransform(frame, el, 'p:xfrm');
		return frame;
	}

	private applyLocalModel3DSize(model: XmlObject, el: Model3DPptxElement): void {
		const shapeProperties = ensureXmlChildOrCreate(model, 'am3d:spPr');
		this.applyModel3DTransform(shapeProperties, { ...el, x: 0, y: 0, xEmu: 0, yEmu: 0 });
	}

	private buildModel3DShapeProperties(el: Model3DPptxElement): XmlObject {
		const result: XmlObject = {
			'a:xfrm': {},
			'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} },
		};
		this.applyModel3DTransform(result, el);
		return result;
	}

	private applyModel3DTransform(parent: XmlObject, el: Model3DPptxElement, tag = 'a:xfrm'): void {
		const transform = ensureXmlChildOrCreate(parent, tag);
		const emu = PptxHandlerRuntime.EMU_PER_PX;
		const width = Math.max(el.width, 1);
		const height = Math.max(el.height, 1);
		transform['a:off'] = {
			'@_x': String(resolveXfrmEmu(el.x, el.xEmu, emu)),
			'@_y': String(resolveXfrmEmu(el.y, el.yEmu, emu)),
		};
		transform['a:ext'] = {
			'@_cx': String(resolveXfrmEmu(width, el.widthEmu, emu)),
			'@_cy': String(resolveXfrmEmu(height, el.heightEmu, emu)),
		};
		this.applyModel3DTransformFlags(transform, el);
	}

	private applyModel3DTransformFlags(transform: XmlObject, el: Model3DPptxElement): void {
		for (const [attribute, value] of [
			['@_rot', el.rotation],
			['@_skewX', el.skewX],
			['@_skewY', el.skewY],
		] as const) {
			if (value === undefined) {
				delete transform[attribute];
			} else {
				transform[attribute] = String(Math.round(value * 60000));
			}
		}
		if (el.flipHorizontal) {
			transform['@_flipH'] = '1';
		} else {
			delete transform['@_flipH'];
		}
		if (el.flipVertical) {
			transform['@_flipV'] = '1';
		} else {
			delete transform['@_flipV'];
		}
	}

	private buildModel3DFallbackPicture(
		el: Model3DPptxElement,
		posterRelationshipId: string | undefined,
	): XmlObject {
		return {
			'p:nvPicPr': {
				'p:cNvPr': { '@_id': el.shapeId ?? '2', '@_name': el.name || el.id },
				'p:cNvPicPr': { 'a:picLocks': { '@_noChangeAspect': '1' } },
				'p:nvPr': {},
			},
			'p:blipFill': {
				'a:blip': posterRelationshipId ? { '@_r:embed': posterRelationshipId } : {},
				'a:stretch': { 'a:fillRect': {} },
			},
			'p:spPr': this.buildModel3DShapeProperties(el),
		};
	}

	private updateModel3DFallback(
		shape: XmlObject,
		el: Model3DPptxElement,
		posterRelationshipId: string,
	): void {
		const block = this.alternateContentBlockByRawXml.get(shape);
		const fallback = block?.rawAc['mc:Fallback'] as XmlObject | undefined;
		const picture =
			(fallback?.['p:pic'] as XmlObject | undefined) ?? this.newModel3DFallbackByXml.get(shape);
		if (!picture) {
			return;
		}
		this.applyModel3DTransform(ensureXmlChildOrCreate(picture, 'p:spPr'), el);
		const fill = ensureXmlChildOrCreate(picture, 'p:blipFill');
		const blip = ensureXmlChildOrCreate(fill, 'a:blip');
		blip['@_r:embed'] = posterRelationshipId;
	}
}
