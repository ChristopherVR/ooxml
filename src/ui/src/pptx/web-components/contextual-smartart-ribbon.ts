import { html, nothing } from 'lit';
import { findDiagramShapeIndex, patchDiagramShape, setDiagramNodeStyle } from 'ooxml-core/diagram';
import type { DiagramNodeStylePatch } from 'ooxml-core/diagram';
import type { PptxSmartArtDrawingShape, SmartArtPptxElement } from 'ooxml-core/pptx';
import {
	smartArtShapeGallery,
	smartArtStyleContext,
	smartArtStylePick,
} from '../render/ribbon-galleries/smartart-format-gallery';
import { buildRibbonGallery } from '../render/ribbon-galleries/gallery-registry';
import type { RibbonGalleryDescriptor } from '../render/ribbon-galleries/gallery-types';
import { shapeGalleryItem } from '../render/shape-gallery';
import type { RibbonGalleryPickEvent } from './ribbon-gallery';
import type { ContextualControls, ContextualFields } from './contextual-ribbon-fields';

export function renderSmartArtControls(
	ctx: ContextualControls,
	fields: ContextualFields,
	element: SmartArtPptxElement,
) {
	const { text, patch, command, number, toggle, color } = fields;
	const data = element.smartArtData;
	const nodes = ctx.nodeIds.flatMap((id) => data?.nodes.find((item) => item.id === id) ?? []);
	const shapes = data?.drawingShapes ?? [];
	const targetShapes = nodes.map(
		(node) => shapes[findDiagramShapeIndex(shapes, data?.nodes ?? [], node.id)],
	);
	const node = nodes[0];
	const shape = targetShapes[0];
	const nodeRequired = nodes.length
		? undefined
		: text('pptx.mediaRibbon.node', 'Select a SmartArt node on the slide.');
	const shapeRequired =
		nodes.length && targetShapes.every(Boolean)
			? undefined
			: text(
					'pptx.mediaRibbon.nodeGeometryUnavailable',
					'The selected node has no editable drawing shape.',
				);
	const updateShape = (update: (shape: PptxSmartArtDrawingShape) => PptxSmartArtDrawingShape) => {
		if (!data || shapeRequired) return;
		let drawingShapes = data.drawingShapes;
		for (const { id } of nodes)
			drawingShapes = patchDiagramShape(drawingShapes, data.nodes, id, update);
		patch({ smartArtData: { ...data, drawingShapes, drawingDirty: true } });
	};
	const style = (stylePatch: DiagramNodeStylePatch) => {
		if (!data || !nodes.length) return;
		const next = nodes.reduce(
			(previous, { id }) => setDiagramNodeStyle(previous, id, stylePatch),
			data,
		);
		patch({ smartArtData: next });
	};
	const gallery = (
		control: string,
		descriptor: RibbonGalleryDescriptor,
		mode: 'inline' | 'dropdown',
		pick: (itemId: string) => void,
	) => {
		const reason = nodeRequired ?? shapeRequired;
		return html`<pptx-ui-ribbon-gallery
			data-ribbon-control=${`${ctx.group}.${control}`}
			mode=${mode}
			.descriptor=${{ ...descriptor, disabled: descriptor.disabled || Boolean(reason), disabledReason: reason }}
			.translateLabel=${ctx.state.translate}
			.disabled=${!fields.available}
			@gallery-pick=${({ detail }: RibbonGalleryPickEvent) => {
				if (fields.available && !reason) pick(detail.itemId);
			}}
		></pptx-ui-ribbon-gallery>`;
	};
	const styleGallery = (id: 'shapeStyles' | 'wordArtStyles') =>
		gallery(
			'gallery',
			buildRibbonGallery(id, smartArtStyleContext(node, shape, data?.pictureScheme)),
			'inline',
			(itemId) => {
				if (!data || shapeRequired) return;
				let next = data;
				for (const [index, node] of nodes.entries()) {
					const result = smartArtStylePick(
						id,
						itemId,
						smartArtStyleContext(node, targetShapes[index], data.pictureScheme),
					);
					if (!result) continue;
					next = setDiagramNodeStyle(next, node.id, result.style);
					next = {
						...next,
						drawingShapes: patchDiagramShape(next.drawingShapes, next.nodes, node.id, (item) => ({
							...item,
							...result.drawing,
						})),
					};
				}
				if (next !== data) patch({ smartArtData: next });
			},
		);
	switch (ctx.group) {
		case 'smartArtFormat.shapes': {
			const resize = (scale: number) =>
				updateShape((item) => {
					const width = Math.max(1, Math.min(9600, item.width * scale));
					const height = Math.max(1, Math.min(9600, item.height * scale));
					return {
						...item,
						width,
						height,
						x: item.x + (item.width - width) / 2,
						y: item.y + (item.height - height) / 2,
					};
				});
			return html`${gallery(
				'shape',
				smartArtShapeGallery(shape ? (shape.shapeType ?? 'rect') : undefined),
				'dropdown',
				(shapeType) => {
					const item = shapeGalleryItem(shapeType);
					if (item?.action.kind !== 'shape' && item?.action.kind !== 'actionButton') return;
					updateShape((current) => ({
						...current,
						shapeType,
						shapeAdjustments: undefined,
						customGeometryRawData: undefined,
						customGeometryPaths: undefined,
						pathData: undefined,
					}));
				},
			)}
			${command('larger', text('pptx.mediaRibbon.larger', 'Larger'), () => resize(1.15), shapeRequired)}
			${command('smaller', text('pptx.mediaRibbon.smaller', 'Smaller'), () => resize(1 / 1.15), shapeRequired)}`;
		}
		case 'smartArtFormat.shapeStyles':
			return html`${styleGallery('shapeStyles')}
			${color('fill', text('pptx.drawing.shapeFill', 'Shape Fill'), node?.style?.fillColor ?? shape?.fillColor ?? '#4472c4', (fillColor) => style({ fillColor }), nodeRequired)}
			${color('border', text('pptx.drawing.shapeOutline', 'Shape Outline'), node?.style?.lineColor ?? shape?.strokeColor ?? '#000000', (lineColor) => style({ lineColor }), nodeRequired)}`;
		case 'smartArtFormat.wordArtStyles':
			return html`${styleGallery('wordArtStyles')}
			${color('color', text('pptx.text.fontColor', 'Font Color'), node?.style?.fontColor ?? shape?.fontColor ?? '#000000', (fontColor) => style({ fontColor }), nodeRequired)}
			${toggle('bold', text('pptx.ribbon.bold', 'Bold'), nodes.length > 0 && nodes.every((node) => node.style?.bold === true), (bold) => style({ bold }), nodeRequired)}
			${toggle('italic', text('pptx.ribbon.italic', 'Italic'), nodes.length > 0 && nodes.every((node) => node.style?.italic === true), (italic) => style({ italic }), nodeRequired)}`;
		case 'smartArtFormat.size':
			return html`${number('width', text('pptx.arrange.width', 'Width'), shape?.width ?? 0, 1, 9600, 'any', (width) => updateShape((item) => ({ ...item, width })), shapeRequired)}
			${number('height', text('pptx.arrange.height', 'Height'), shape?.height ?? 0, 1, 9600, 'any', (height) => updateShape((item) => ({ ...item, height })), shapeRequired)}`;
		default:
			return nothing;
	}
}
