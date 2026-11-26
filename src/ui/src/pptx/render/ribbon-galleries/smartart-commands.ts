/**
 * SmartArt Design > Create Graphic and Reset: PowerPoint's command buttons,
 * expressed as one-item gallery modules so every binding renders and dispatches
 * them through the gallery path it already has (descriptor, `gallery-pick`,
 * element patch / replace / host command + history).
 *
 * - Add Shape / Add Bullet / Add Shape Before / Add Shape Above run the core
 *   node edits (`addSmartArtNode`, `addSmartArtNodeAsChild`,
 *   `addSmartArtNodeBefore`, `addSmartArtNodeAbove`) on the selected node when
 *   known, or the boundary top-level node otherwise.
 * - Text Pane requests the host command (`smartArtTextPane`) to toggle/open the
 *   Inspector's SmartArt text pane.
 * - Promote, Demote, Move Up and Move Down run `promoteSmartArtNode`,
 *   `demoteSmartArtNode` and `reorderSmartArtNode` on the selected node (or,
 *   without an explicit node selection, the last/first node for which the edit
 *   is valid) and disable themselves when no valid edit exists.
 * - Right to Left toggles `presLayoutVars.direction` (`'rev'` <-> `'norm'`) and
 *   rebuilds drawing shapes so the layout flips direction.
 * - Reset Graphic clears direct node formatting while retaining diagram choices.
 * - Convert decomposes the SmartArt into standard `PptxElement[]` shapes via
 *   `decomposeSmartArt` and returns a `replaceElement` result.
 *
 * @module render/ribbon-galleries/smartart-commands
 */
import type { PptxElement, PptxSmartArtData } from 'ooxml-core/pptx';
import {
	addSmartArtNode,
	addSmartArtNodeAbove,
	addSmartArtNodeAsChild,
	addSmartArtNodeBefore,
	decomposeSmartArt,
	demoteSmartArtNode,
	promoteSmartArtNode,
	reorderSmartArtNode,
	resetSmartArtFormatting,
} from 'ooxml-core/pptx';

import type { RibbonGalleryModule } from './gallery-module';
import type {
	RibbonGalleryApplyResult,
	RibbonGalleryContext,
	RibbonGalleryId,
} from './gallery-types';
import { smartArtElementPatch } from './smartart-gallery-patch';

interface CommandSpec {
	id: RibbonGalleryId;
	labelKey: string;
	label: string;
	iconPath: string;
	large?: boolean;
	/** Return true when the command cannot run on this SmartArt right now. */
	isDisabled?: (data: PptxSmartArtData, ctx: RibbonGalleryContext) => boolean;
	/** Optional tooltip explaining why the command is disabled. */
	disabledHint?: { key: string; text: string };
	run: (data: PptxSmartArtData, ctx: RibbonGalleryContext) => RibbonGalleryApplyResult | null;
}

const NEEDS_NODE = {
	key: 'pptx.gallery.smartArtCommand.needsNode',
	text: 'The SmartArt hierarchy does not allow this command for the current selection.',
};

/** Gallery-mode commands whose target follows the canvas selection in the contextual strip. */
export const SMARTART_NODE_COMMANDS = [
	['promote', 'smartArtPromote'],
	['demote', 'smartArtDemote'],
	['moveUp', 'smartArtMoveUp'],
	['moveDown', 'smartArtMoveDown'],
] as const;

function topLevelIds(data: PptxSmartArtData): string[] {
	const ids = new Set(data.nodes.map((node) => node.id));
	return data.nodes
		.filter((node) => !node.parentId || !ids.has(node.parentId))
		.map((node) => node.id);
}

function lastTopLevelId(data: PptxSmartArtData): string | undefined {
	return topLevelIds(data).at(-1);
}

function selectedNodeId(data: PptxSmartArtData, ctx: RibbonGalleryContext): string | undefined {
	const id = ctx.selectedSmartArtNodeId;
	if (!id) {
		return undefined;
	}
	return data.nodes.some((node) => node.id === id) ? id : undefined;
}

function resolveNodeOpTarget(
	data: PptxSmartArtData,
	ctx: RibbonGalleryContext,
	op: (d: PptxSmartArtData, id: string) => PptxSmartArtData,
	searchFromEnd: boolean,
): string | undefined {
	const explicit = selectedNodeId(data, ctx);
	if (explicit) {
		return op(data, explicit) !== data ? explicit : undefined;
	}
	const nodes = data.nodes;
	if (searchFromEnd) {
		for (let i = nodes.length - 1; i >= 0; i -= 1) {
			const node = nodes[i];
			if (node && op(data, node.id) !== data) {
				return node.id;
			}
		}
		return undefined;
	}
	for (const node of nodes) {
		if (op(data, node.id) !== data) {
			return node.id;
		}
	}
	return undefined;
}

function viaData(
	element: PptxElement | null,
	next: PptxSmartArtData | undefined,
	current: PptxSmartArtData,
): RibbonGalleryApplyResult | null {
	if (!next || next === current) {
		return null;
	}
	const patch = smartArtElementPatch(element, next);
	return patch ? { kind: 'element', ...patch } : null;
}

const SPECS: readonly CommandSpec[] = [
	{
		id: 'smartArtAddShape',
		labelKey: 'pptx.gallery.smartArtCommand.addShape',
		label: 'Add Shape',
		iconPath: 'M3 4h14v12H3zM10 7v6M7 10h6',
		run: (data, ctx) =>
			viaData(
				ctx.element,
				addSmartArtNode(data, '', selectedNodeId(data, ctx) ?? lastTopLevelId(data)),
				data,
			),
	},
	{
		id: 'smartArtAddShapeBefore',
		labelKey: 'pptx.gallery.smartArtCommand.addShapeBefore',
		label: 'Add Shape Before',
		iconPath: 'M8 4h9v12H8zM4 10h6M4 10l2-2M4 10l2 2',
		run: (data, ctx) =>
			viaData(
				ctx.element,
				addSmartArtNodeBefore(data, '', selectedNodeId(data, ctx) ?? topLevelIds(data)[0]),
				data,
			),
	},
	{
		id: 'smartArtAddShapeAbove',
		labelKey: 'pptx.gallery.smartArtCommand.addShapeAbove',
		label: 'Add Shape Above',
		iconPath: 'M5 3h10v5H5zM5 12h10v5H5zM10 8v4',
		run: (data, ctx) => {
			const target = selectedNodeId(data, ctx) ?? lastTopLevelId(data);
			return target ? viaData(ctx.element, addSmartArtNodeAbove(data, target, ''), data) : null;
		},
	},
	{
		id: 'smartArtAddBullet',
		labelKey: 'pptx.gallery.smartArtCommand.addBullet',
		label: 'Add Bullet',
		iconPath: 'M4 6h.01M8 6h8M4 10h.01M8 10h8M4 14h.01M8 14h8',
		run: (data, ctx) =>
			viaData(
				ctx.element,
				addSmartArtNodeAsChild(data, selectedNodeId(data, ctx) ?? lastTopLevelId(data), ''),
				data,
			),
	},
	{
		id: 'smartArtTextPane',
		labelKey: 'pptx.gallery.smartArtCommand.textPane',
		label: 'Text Pane',
		iconPath: 'M3 3.5h14v13H3zM7 3.5v13M10 7h5M10 10h5',
		run: () => ({ kind: 'command', command: 'smartArtTextPane' }),
	},
	{
		id: 'smartArtPromote',
		labelKey: 'pptx.gallery.smartArtCommand.promote',
		label: 'Promote',
		iconPath: 'M16 10H5M8 6.5 4.5 10 8 13.5',
		disabledHint: NEEDS_NODE,
		isDisabled: (data, ctx) => !resolveNodeOpTarget(data, ctx, promoteSmartArtNode, true),
		run: (data, ctx) => {
			const target = resolveNodeOpTarget(data, ctx, promoteSmartArtNode, true);
			return target ? viaData(ctx.element, promoteSmartArtNode(data, target), data) : null;
		},
	},
	{
		id: 'smartArtDemote',
		labelKey: 'pptx.gallery.smartArtCommand.demote',
		label: 'Demote',
		iconPath: 'M4 10h11M12 6.5l3.5 3.5-3.5 3.5',
		disabledHint: NEEDS_NODE,
		isDisabled: (data, ctx) => !resolveNodeOpTarget(data, ctx, demoteSmartArtNode, true),
		run: (data, ctx) => {
			const target = resolveNodeOpTarget(data, ctx, demoteSmartArtNode, true);
			return target ? viaData(ctx.element, demoteSmartArtNode(data, target), data) : null;
		},
	},
	{
		id: 'smartArtRightToLeft',
		labelKey: 'pptx.gallery.smartArtCommand.rightToLeft',
		label: 'Right to Left',
		iconPath: 'M4 7h12M13 4l3 3-3 3M16 13H4M7 10l-3 3 3 3',
		run: (data, ctx) => {
			const nextDirection = data.presLayoutVars?.direction === 'rev' ? 'norm' : 'rev';
			const next: PptxSmartArtData = {
				...data,
				presLayoutVars: { ...data.presLayoutVars, direction: nextDirection },
				...(data.drawingShapes && data.drawingShapes.length > 0 ? { drawingShapes: [] } : {}),
				drawingDirty: true,
			};
			return viaData(ctx.element, next, data);
		},
	},
	{
		id: 'smartArtMoveUp',
		labelKey: 'pptx.gallery.smartArtCommand.moveUp',
		label: 'Move Up',
		iconPath: 'M10 16V5M6.5 8 10 4.5 13.5 8',
		disabledHint: NEEDS_NODE,
		isDisabled: (data, ctx) =>
			!resolveNodeOpTarget(data, ctx, (d, id) => reorderSmartArtNode(d, id, -1), true),
		run: (data, ctx) => {
			const target = resolveNodeOpTarget(
				data,
				ctx,
				(d, id) => reorderSmartArtNode(d, id, -1),
				true,
			);
			return target ? viaData(ctx.element, reorderSmartArtNode(data, target, -1), data) : null;
		},
	},
	{
		id: 'smartArtMoveDown',
		labelKey: 'pptx.gallery.smartArtCommand.moveDown',
		label: 'Move Down',
		iconPath: 'M10 4v11M6.5 12 10 15.5 13.5 12',
		disabledHint: NEEDS_NODE,
		isDisabled: (data, ctx) =>
			!resolveNodeOpTarget(data, ctx, (d, id) => reorderSmartArtNode(d, id, 1), false),
		run: (data, ctx) => {
			const target = resolveNodeOpTarget(
				data,
				ctx,
				(d, id) => reorderSmartArtNode(d, id, 1),
				false,
			);
			return target ? viaData(ctx.element, reorderSmartArtNode(data, target, 1), data) : null;
		},
	},
	{
		id: 'smartArtResetGraphic',
		labelKey: 'pptx.gallery.smartArtCommand.resetGraphic',
		label: 'Reset Graphic',
		iconPath: 'M4 10a6 6 0 1 0 2-4.5M4 3.5V7h3.5',
		large: true,
		run: (data, ctx) => {
			const element = ctx.element;
			if (element?.type !== 'smartArt') return null;
			return {
				kind: 'element',
				elementId: element.id,
				patch: {
					smartArtData: resetSmartArtFormatting(data, element, ctx.themeColorMap),
				},
			};
		},
	},
	{
		id: 'smartArtConvert',
		labelKey: 'pptx.gallery.smartArtCommand.convert',
		label: 'Convert',
		iconPath: 'M4 6h8l-2-2M16 14H8l2 2M4 6v3M16 14v-3',
		large: true,
		isDisabled: (data) => data.nodes.length === 0,
		run: (data, ctx) => {
			const element = ctx.element;
			if (element?.type !== 'smartArt') {
				return null;
			}
			const elements = decomposeSmartArt(
				data,
				{ x: element.x, y: element.y, width: element.width, height: element.height },
				ctx.themeColorMap ? { ...ctx.themeColorMap } : undefined,
			);
			if (!elements || elements.length === 0) {
				return null;
			}
			return { kind: 'replaceElement', elementId: element.id, elements };
		},
	},
];

function moduleFor(spec: CommandSpec): RibbonGalleryModule {
	return {
		build(ctx) {
			const data = ctx.element?.type === 'smartArt' ? ctx.element.smartArtData : undefined;
			const disabled = !data || Boolean(spec.isDisabled?.(data, ctx));
			const hint = disabled && data ? spec.disabledHint : undefined;
			return {
				id: spec.id,
				labelKey: spec.labelKey,
				label: spec.label,
				disabled,
				command: {
					iconPath: spec.iconPath,
					large: spec.large,
					hintKey: hint?.key,
					hint: hint?.text,
				},
				sections: [
					{
						id: 'command',
						columns: 1,
						tileWidth: 0,
						tileHeight: 0,
						items: [
							{
								id: 'run',
								labelKey: spec.labelKey,
								label: spec.label,
								previewSvg: '',
								applied: false,
							},
						],
					},
				],
			};
		},
		apply(itemId, ctx) {
			const data = ctx.element?.type === 'smartArt' ? ctx.element.smartArtData : undefined;
			if (itemId !== 'run' || !data || spec.isDisabled?.(data, ctx)) {
				return null;
			}
			return spec.run(data, ctx);
		},
	};
}

/** Gallery ids that are SmartArt command buttons. */
export type SmartArtCommandGalleryId = Extract<
	RibbonGalleryId,
	| 'smartArtAddShape'
	| 'smartArtAddShapeBefore'
	| 'smartArtAddShapeAbove'
	| 'smartArtAddBullet'
	| 'smartArtTextPane'
	| 'smartArtPromote'
	| 'smartArtDemote'
	| 'smartArtRightToLeft'
	| 'smartArtMoveUp'
	| 'smartArtMoveDown'
	| 'smartArtResetGraphic'
	| 'smartArtConvert'
>;

/** The SmartArt command entries, keyed by gallery id, for the registry. */
export const SMARTART_COMMAND_MODULES = Object.fromEntries(
	SPECS.map((spec) => [spec.id, moduleFor(spec)]),
) as Record<SmartArtCommandGalleryId, RibbonGalleryModule>;
