import { html, nothing } from 'lit';
import { findDiagramShapeIndex, patchDiagramShape, setDiagramNodeStyle } from 'ooxml-core/diagram';
import type { DiagramNodeStylePatch } from 'ooxml-core/diagram';
import type { PptxSmartArtDrawingShape, SmartArtPptxElement } from 'ooxml-core/pptx';
import type { ContextualControls, ContextualFields } from './contextual-ribbon-fields';

export function renderSmartArtControls(
	ctx: ContextualControls,
	fields: ContextualFields,
	element: SmartArtPptxElement,
) {
	const { text, patch, select, command, number, toggle, color } = fields;
	const data = element.smartArtData;
	const nodes = data?.nodes.slice(0, 512) ?? [];
	const node = nodes.find((item) => item.id === ctx.nodeId) ?? nodes[0];
	const shapes = data?.drawingShapes;
	const shape = node && shapes?.[findDiagramShapeIndex(shapes, nodes, node.id)];
	const nodeRequired = node ? undefined : fields.unavailable;
	const shapeRequired = shape
		? undefined
		: text(
				'pptx.mediaRibbon.nodeGeometryUnavailable',
				'The selected node has no editable drawing shape.',
			);
	const updateShape = (update: (shape: PptxSmartArtDrawingShape) => PptxSmartArtDrawingShape) => {
		if (!data || !node || !shape) return;
		const drawingShapes = patchDiagramShape(shapes, data.nodes, node.id, update);
		patch({ smartArtData: { ...data, drawingShapes, drawingDirty: true } });
	};
	const style = (stylePatch: DiagramNodeStylePatch) => {
		if (data && node) patch({ smartArtData: setDiagramNodeStyle(data, node.id, stylePatch) });
	};
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
			return html`${select(
				'node',
				text('pptx.mediaRibbon.node', 'SmartArt Node'),
				node?.id ?? '',
				nodes.map((item, index) => [item.id, item.text || String(index + 1)]),
				ctx.selectNode,
				nodeRequired,
			)}
			${select(
				'shape',
				text('pptx.mediaRibbon.changeShape', 'Change Shape'),
				shape?.shapeType ?? '',
				[
					['rect', text('pptx.ribbon.rectangle', 'Rectangle')],
					['ellipse', text('pptx.ribbon.ellipse', 'Ellipse')],
				],
				(shapeType) =>
					updateShape((item) => ({
						...item,
						shapeType,
						customGeometryRawData: undefined,
						pathData: undefined,
					})),
				shapeRequired,
			)}
			${command('larger', text('pptx.mediaRibbon.larger', 'Larger'), () => resize(1.15), shapeRequired)}
			${command('smaller', text('pptx.mediaRibbon.smaller', 'Smaller'), () => resize(1 / 1.15), shapeRequired)}`;
		}
		case 'smartArtFormat.shapeStyles':
			return html`${color('fill', text('pptx.drawing.shapeFill', 'Shape Fill'), node?.style?.fillColor ?? shape?.fillColor ?? '#4472c4', (fillColor) => style({ fillColor }), nodeRequired)}
			${color('border', text('pptx.drawing.shapeOutline', 'Shape Outline'), node?.style?.lineColor ?? shape?.strokeColor ?? '#000000', (lineColor) => style({ lineColor }), nodeRequired)}`;
		case 'smartArtFormat.wordArtStyles':
			return html`${color('color', text('pptx.text.fontColor', 'Font Color'), node?.style?.fontColor ?? shape?.fontColor ?? '#000000', (fontColor) => style({ fontColor }), nodeRequired)}
			${toggle('bold', text('pptx.ribbon.bold', 'Bold'), node?.style?.bold === true, (bold) => style({ bold }), nodeRequired)}
			${toggle('italic', text('pptx.ribbon.italic', 'Italic'), node?.style?.italic === true, (italic) => style({ italic }), nodeRequired)}`;
		case 'smartArtFormat.arrange':
			return html`${number('x', text('pptx.mediaRibbon.x', 'Horizontal Position'), shape?.x ?? 0, -9600, 9600, 'any', (x) => updateShape((item) => ({ ...item, x })), shapeRequired)}
			${number('y', text('pptx.mediaRibbon.y', 'Vertical Position'), shape?.y ?? 0, -9600, 9600, 'any', (y) => updateShape((item) => ({ ...item, y })), shapeRequired)}`;
		case 'smartArtFormat.size':
			return html`${number('width', text('pptx.arrange.width', 'Width'), shape?.width ?? 0, 1, 9600, 'any', (width) => updateShape((item) => ({ ...item, width })), shapeRequired)}
			${number('height', text('pptx.arrange.height', 'Height'), shape?.height ?? 0, 1, 9600, 'any', (height) => updateShape((item) => ({ ...item, height })), shapeRequired)}`;
		default:
			return nothing;
	}
}
