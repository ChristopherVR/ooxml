// Package-independent preview of an uncached diagram. Keep the original definitions intact:
// the per-point engine reads their ordered constraints and conditional layout statements.
import { parseXml } from '../xml/index';
import { resolveDrawingColor } from '../drawingml/drawing-color';
import { descendants, first, NS } from './dom';
import { parseDiagramDataModel } from './data-model';
import { parseDiagramColors, parseDiagramQuickStyle } from './definitions';
import { runEngineLayout } from './engine';
import { applyCustomLayoutOverrides } from './layout/smartart-layout-interpreter-custom';
import { diagramDrawingFromViews } from './drawing-from-views';
import { diagramPreviewText } from './layout-text';
import { approximateDiagramDrawingFromNodes, type ApproximateDiagramInput } from './approximate';
import type { DiagramData, DiagramNode, DiagramPresLayoutVars } from './model';
import type { DiagramHostNode } from './edit/node-edit';
import type { DiagramDrawing, DiagramIssue } from './types';

/** Original part text, JSON-safe so hosts can retain it through undo and collaboration. */
export interface DiagramLayoutSource {
	layoutXml?: string;
	dataXml?: string;
	colorsXml?: string;
	quickStyleXml?: string;
}

export interface DiagramPreviewInput extends ApproximateDiagramInput {
	nodes: readonly DiagramHostNode[];
	source?: DiagramLayoutSource | undefined;
	scheme?: Readonly<Record<string, string>> | undefined;
	minorFont?: string | undefined;
}

export interface DiagramPreview {
	drawing?: DiagramDrawing;
	mode: 'interpreted' | 'approximate' | 'placeholder';
	issues: DiagramIssue[];
}

/** Adapt current host nodes, keeping source metadata but never restoring stale text or order. */
export function diagramLayoutData(input: DiagramPreviewInput): DiagramData {
	const original = input.source?.dataXml ? parseDiagramDataModel(input.source.dataXml) : undefined;
	const nodes: DiagramNode[] = input.nodes.map((node) => {
		const point = original?.nodes.find((entry) => entry.modelId === node.id);
		const presentation = original?.points.find(
			(entry) => entry.type === 'pres' && entry.presAssocId === node.id,
		);
		const customLayout = point?.customLayout ?? presentation?.customLayout;
		const styleRole = point?.presStyleLabel ?? presentation?.presStyleLabel;
		return {
			id: node.id,
			text: node.text,
			...(node.parentId ? { parentId: node.parentId } : {}),
			...(node.nodeType ? { nodeType: node.nodeType } : {}),
			...(node.paragraphs ? { textParagraphs: node.paragraphs } : {}),
			...(customLayout ? { customLayout } : {}),
			...(styleRole ? { styleRole } : {}),
		};
	});
	// Source connections retain transition labels, while current host order and parents win.
	const connections = nodes.map((node, index) => {
		const connection = original?.connections.find(
			(entry) => entry.destId === node.id && (!entry.type || entry.type === 'parOf'),
		);
		return {
			...connection,
			modelId: connection?.modelId ?? `preview-${node.id}`,
			type: 'parOf',
			sourceId: node.parentId ?? '__doc__',
			destId: node.id,
			srcOrd: index,
		};
	});
	const data: DiagramData = { nodes, connections };
	if (input.minorFont) data.themeMinorFont = input.minorFont;
	if (input.family) data.resolvedLayoutType = input.family;
	if (input.source?.layoutXml)
		data.layoutDefinition = { rootNode: {}, rawXmlText: input.source.layoutXml };
	if (input.source?.dataXml) {
		const root = parseXml(input.source.dataXml).documentElement;
		const vars = descendants(root, 'presLayoutVars')[0];
		const direction = first(vars, 'dir', NS.dgm)?.getAttribute('val');
		const values: DiagramPresLayoutVars = {};
		if (direction === 'norm' || direction === 'rev') values.direction = direction;
		const branch = first(vars, 'hierBranch', NS.dgm)?.getAttribute('val');
		if (
			branch === 'std' ||
			branch === 'init' ||
			branch === 'l' ||
			branch === 'r' ||
			branch === 'hang'
		)
			values.hierarchyBranch = branch;
		for (const [tag, key] of [
			['chMax', 'childMax'],
			['chPref', 'childPreferred'],
		] as const) {
			const raw = first(vars, tag, NS.dgm)?.getAttribute('val');
			if (raw && Number.isFinite(Number(raw))) values[key] = Number(raw);
		}
		for (const [tag, key] of [
			['orgChart', 'orgChart'],
			['bulletEnabled', 'bulletEnabled'],
		] as const) {
			const raw = first(vars, tag, NS.dgm)?.getAttribute('val');
			if (raw) values[key] = raw === '1' || raw === 'true';
		}
		for (const [tag, key] of [
			['animLvl', 'animationLevel'],
			['animOne', 'animateOne'],
			['resizeHandles', 'resizeHandles'],
		] as const) {
			const raw = first(vars, tag, NS.dgm)?.getAttribute('val');
			if (raw) values[key] = raw;
		}
		data.presLayoutVars = values;
	}
	if (input.source?.colorsXml) {
		const colors = parseDiagramColors(input.source.colorsXml);
		const roleColors: NonNullable<NonNullable<DiagramData['colorTransform']>['roleColors']> = {};
		for (const label of colors.labels) {
			const fill = label.fill.colors.flatMap((color) => {
				const resolved = resolveDrawingColor(color, { scheme: (name) => input.scheme?.[name] });
				return resolved ? [resolved.hex] : [];
			});
			if (fill.length) roleColors[label.name] = { fill, line: [] };
		}
		data.colorTransform = { roleColors, fillColors: [], lineColors: [] };
	}
	return data;
}

/** Try the shared ordered-layout engine, then an explicitly reported family approximation. */
export function previewDiagramLayout(input: DiagramPreviewInput): DiagramPreview {
	const issues: DiagramIssue[] = [];
	try {
		const data = diagramLayoutData(input);
		const palette = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map(
			(name) => input.scheme?.[name] ?? '#4472c4',
		);
		const box = { width: input.extentEmu.width / 9525, height: input.extentEmu.height / 9525 };
		const computed = runEngineLayout(data, box, data.nodes, palette, 'flat');
		if (computed) {
			const result = applyCustomLayoutOverrides(computed, data.nodes, box);
			const views = result.nodes.flatMap((node) => {
				if (node.kind !== 'rect') return [];
				return [
					{
						id: node.key,
						x: node.x,
						y: node.y,
						width: node.width,
						height: node.height,
						shapeType: node.presetOverride ?? 'rect',
						fillNone: node.fill === 'none',
						fillColor: node.fill,
						strokeColor: node.stroke,
						strokeWidth: node.strokeWidth,
						text:
							node.literalText ??
							[
								node.text,
								...(node.foldedNodeIds ?? []).map(
									(id) => input.nodes.find((item) => item.id === id)?.text ?? '',
								),
							]
								.filter(Boolean)
								.join('\n'),
						fontSize: node.fontSize,
						fillOpacity: node.opacity,
						...(node.rotation ? { rotation: node.rotation } : {}),
						...(node.shapeAdjustments ? { shapeAdjustments: node.shapeAdjustments } : {}),
					},
				];
			});
			const drawing = diagramDrawingFromViews(views);
			const quickStyle = input.source?.quickStyleXml
				? parseDiagramQuickStyle(input.source.quickStyleXml)
				: undefined;
			result.nodes.forEach((node, index) => {
				const shape = drawing.shapes[index];
				if (!shape || node.kind !== 'rect' || node.literalText !== undefined) return;
				const label = quickStyle?.labels.find((entry) => entry.name === node.styleLabel);
				if (label) {
					shape.style = label;
					shape.has3d = label.hasScene3d || label.hasShape3d;
				}
				const sourceIds = [node.nodeId, ...(node.foldedNodeIds ?? [])];
				const textNodes = sourceIds.flatMap((id) => {
					const source = input.nodes.find((entry) => entry.id === id);
					return source ? [source] : [];
				});
				if (textNodes.length) shape.text = diagramPreviewText(textNodes, node.fontSize);
			});
			return { drawing, mode: 'interpreted', issues };
		}
	} catch (error) {
		issues.push({
			code: 'DIAGRAM_LAYOUT_FAILED',
			message: error instanceof Error ? error.message : String(error),
		});
	}
	issues.push({
		code: 'DIAGRAM_LAYOUT_APPROXIMATE',
		message: 'The layout definition could not be interpreted; a family approximation is used.',
	});
	const drawing = approximateDiagramDrawingFromNodes(input);
	return { ...(drawing ? { drawing } : {}), mode: drawing ? 'approximate' : 'placeholder', issues };
}
