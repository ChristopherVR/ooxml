import { EMU_PER_PIXEL } from '../../units/constants';
import { NS, elements, parseXml, type XmlElement } from '../../xml/index';
import type { DrawingObject, Worksheet } from '../model';
import type { ShapeNode } from '../model-shapes';
import { anchorToPixelBox, pixelBoxToAnchor, type PixelBox } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import { selfContainedXml } from '../read/xml-util';
import { patchWordArtShape } from '../word-art';
import { patchNativeShape } from '../write/native-shape';
import { type EditContext, sheetAt } from './context';
import { editDrawings, isWorksheetShape } from './native-shapes';

const leafTags = new Set(['sp', 'cxnSp']);
function hasLinkedContent(node: XmlElement): boolean {
	return (
		(node.namespaceURI === NS.a && ['stCxn', 'endCxn'].includes(node.localName)) ||
		Array.from(node.attributes).some((attr) => attr.namespaceURI === NS.r) ||
		elements(node).some(hasLinkedContent)
	);
}

function selfContained(xml: string): boolean {
	return !hasLinkedContent(parseXml(xml).documentElement);
}

const unique = (indices: readonly number[]) => [...new Set(indices)].sort((a, b) => a - b);

export function canGroupDrawings(sheet: Worksheet, indices: readonly number[]): boolean {
	return (
		unique(indices).length > 1 &&
		indices.every((index) => {
			const drawing = sheet.drawings[index];
			return (
				isWorksheetShape(drawing) &&
				drawing.shape.kind !== 'group' &&
				!drawing.shape.equationParagraphs &&
				(!drawing.sourceXml || selfContained(drawing.sourceXml))
			);
		})
	);
}

function shapeContent(drawing: DrawingObject | undefined): ShapeNode | undefined {
	if (!isWorksheetShape(drawing)) return undefined;
	const shape = structuredClone(drawing.shape);
	if (!drawing.sourceXml) return shape;
	const source = drawing.wordArtEdited
		? patchWordArtShape(drawing.sourceXml, shape)
		: drawing.sourceXml;
	const xml = drawing.shapeEdit ? patchNativeShape(source, shape, drawing.shapeEdit) : source;
	const root = parseXml(xml).documentElement;
	const node = elements(root).find(
		(child) => child.namespaceURI === NS.xdr && leafTags.has(child.localName),
	);
	if (!node) return undefined;
	shape.sourceShapeXml = selfContainedXml(node);
	return shape;
}

function emuFrame(box: PixelBox): ShapeNode['frame'] {
	return {
		x: Math.round(box.x * EMU_PER_PIXEL),
		y: Math.round(box.y * EMU_PER_PIXEL),
		width: Math.round(box.w * EMU_PER_PIXEL),
		height: Math.round(box.h * EMU_PER_PIXEL),
	};
}

export function groupDrawings(
	ctx: EditContext,
	s: number,
	indices: readonly number[],
	name: string,
): number | undefined {
	const sheet = sheetAt(ctx.workbook, s);
	if (!canGroupDrawings(sheet, indices)) return undefined;
	const sorted = unique(indices);
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const items = sorted.flatMap((index) => {
		const drawing = sheet.drawings[index];
		const shape = shapeContent(drawing);
		return shape && drawing
			? [{ shape, box: anchorToPixelBox(sheet, drawing.anchor, metrics) }]
			: [];
	});
	if (items.length !== sorted.length) return undefined;
	const x = Math.min(...items.map(({ box }) => box.x));
	const y = Math.min(...items.map(({ box }) => box.y));
	const w = Math.max(...items.map(({ box }) => box.x + box.w)) - x;
	const h = Math.max(...items.map(({ box }) => box.y + box.h)) - y;
	const frame = emuFrame({ x: 0, y: 0, w, h });
	const group: DrawingObject = {
		kind: 'unsupported',
		description: 'group',
		anchor: pixelBoxToAnchor(sheet, { x, y, w, h }, metrics),
		shape: {
			kind: 'group',
			geometry: 'group',
			name,
			frame,
			childFrame: { ...frame },
			children: items.map(({ shape, box }) => ({
				...shape,
				frame: emuFrame({ ...box, x: box.x - x, y: box.y - y }),
			})),
		},
	};
	const top = sorted.at(-1);
	return editDrawings(ctx, s, 'Group objects', () => {
		const selected = new Set(sorted);
		sheet.drawings = sheet.drawings.flatMap((drawing, index) =>
			index === top ? [group] : selected.has(index) ? [] : [drawing],
		);
		return sheet.drawings.indexOf(group);
	});
}

function groupChildren(drawing: DrawingObject | undefined): ShapeNode[] | undefined {
	if (!isWorksheetShape(drawing)) return undefined;
	const { shape } = drawing;
	if (
		shape.kind !== 'group' ||
		!shape.childFrame?.width ||
		!shape.childFrame.height ||
		shape.flipH ||
		shape.flipV ||
		!shape.children?.length ||
		shape.children.some((child) => child.kind === 'group')
	)
		return undefined;
	if (!drawing.sourceXml) return shape.children;
	if (!selfContained(drawing.sourceXml)) return undefined;
	const group = elements(parseXml(drawing.sourceXml).documentElement).find(
		(child) => child.namespaceURI === NS.xdr && child.localName === 'grpSp',
	);
	const children =
		group &&
		elements(group).filter(
			(child) =>
				child.namespaceURI === NS.xdr && !['nvGrpSpPr', 'grpSpPr'].includes(child.localName),
		);
	if (
		!children ||
		children.length !== shape.children.length ||
		children.some((child) => !leafTags.has(child.localName))
	)
		return undefined;
	return children.flatMap((child, index) => {
		const item = shape.children?.[index];
		return item ? [{ ...item, sourceShapeXml: selfContainedXml(child) }] : [];
	});
}

export function canUngroupDrawing(drawing: DrawingObject | undefined): boolean {
	return groupChildren(drawing) !== undefined;
}

export function ungroupDrawing(ctx: EditContext, s: number, index: number): number[] {
	const sheet = sheetAt(ctx.workbook, s);
	const drawing = sheet.drawings[index];
	const children = groupChildren(drawing);
	if (!children || !isWorksheetShape(drawing) || !drawing.shape.childFrame) return [];
	const childFrame = drawing.shape.childFrame;
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const box = anchorToPixelBox(sheet, drawing.anchor, metrics);
	const degrees = drawing.shape.rotation ?? 0;
	const angle = (degrees * Math.PI) / 180;
	const cos = Math.cos(angle);
	const sin = Math.sin(angle);
	const objects = children.map((shape): DrawingObject => {
		const w = (shape.frame.width / childFrame.width) * box.w;
		const h = (shape.frame.height / childFrame.height) * box.h;
		const dx = ((shape.frame.x - childFrame.x) / childFrame.width) * box.w + w / 2 - box.w / 2;
		const dy = ((shape.frame.y - childFrame.y) / childFrame.height) * box.h + h / 2 - box.h / 2;
		const x = Math.max(0, box.x + box.w / 2 + dx * cos - dy * sin - w / 2);
		const y = Math.max(0, box.y + box.h / 2 + dx * sin + dy * cos - h / 2);
		return {
			kind: 'unsupported',
			description: shape.textBox ? 'text box' : 'shape',
			anchor: pixelBoxToAnchor(sheet, { x, y, w, h }, metrics),
			shape: {
				...shape,
				rotation: ((shape.rotation ?? 0) + degrees) % 360,
				frame: emuFrame({ x: 0, y: 0, w, h }),
			},
		};
	});
	return editDrawings(ctx, s, 'Ungroup objects', () => {
		sheet.drawings.splice(index, 1, ...objects);
		return objects.map((_, offset) => index + offset);
	});
}
