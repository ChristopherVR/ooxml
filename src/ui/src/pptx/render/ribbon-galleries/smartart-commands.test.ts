import type { PptxElement, PptxSmartArtData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { replaceElementInList } from './gallery-element-patch';
import { applyRibbonGalleryItem, buildRibbonGallery } from './gallery-registry';
import type { RibbonGalleryContext, RibbonGalleryId } from './gallery-types';

const smartArt = {
	id: 'sa1',
	type: 'smartArt',
	x: 10,
	y: 20,
	width: 400,
	height: 200,
	smartArtData: {
		layout: 'basicBlockList',
		resolvedLayoutType: 'list',
		nodes: [
			{ id: 'n1', text: 'One' },
			{ id: 'n2', text: 'Two' },
		],
		connections: [],
	},
} as unknown as PptxElement;

const hierarchicalSmartArt = {
	id: 'sa2',
	type: 'smartArt',
	x: 0,
	y: 0,
	width: 400,
	height: 200,
	smartArtData: {
		layout: 'basicBlockList',
		resolvedLayoutType: 'hierarchy',
		nodes: [
			{ id: 'n1', text: 'Parent' },
			{ id: 'n1a', text: 'Child', parentId: 'n1' },
			{ id: 'n2', text: 'Sibling' },
		],
		connections: [{ sourceId: 'n1', destId: 'n1a', type: 'parOf' }],
	},
} as unknown as PptxElement;

const singleNodeSmartArt = {
	id: 'sa3',
	type: 'smartArt',
	x: 0,
	y: 0,
	width: 400,
	height: 200,
	smartArtData: {
		layout: 'basicBlockList',
		resolvedLayoutType: 'list',
		nodes: [{ id: 'n1', text: 'Only' }],
		connections: [],
	},
} as unknown as PptxElement;

const ctx: RibbonGalleryContext = { element: smartArt };
const run = (id: RibbonGalleryId, context: RibbonGalleryContext = ctx) =>
	applyRibbonGalleryItem(id, 'run', context);

describe('smartArt Design commands', () => {
	it('draws each command as one enabled, command-style entry on a SmartArt', () => {
		for (const id of [
			'smartArtAddShape',
			'smartArtAddShapeBefore',
			'smartArtAddShapeAbove',
			'smartArtAddBullet',
			'smartArtTextPane',
			'smartArtDemote',
			'smartArtRightToLeft',
			'smartArtMoveUp',
			'smartArtMoveDown',
			'smartArtResetGraphic',
			'smartArtConvert',
		] as const) {
			const descriptor = buildRibbonGallery(id, ctx);
			expect(descriptor.disabled).toBeFalsy();
			expect(descriptor.command?.iconPath).toMatch(/^M/u);
			expect(descriptor.sections[0].items).toHaveLength(1);
		}
		expect(
			buildRibbonGallery('smartArtPromote', { element: hierarchicalSmartArt }).disabled,
		).toBeFalsy();
		expect(buildRibbonGallery('smartArtResetGraphic', ctx).command?.large).toBeTruthy();
		expect(buildRibbonGallery('smartArtConvert', ctx).command?.large).toBeTruthy();
	});

	it('disables every command without a SmartArt selection', () => {
		for (const id of [
			'smartArtAddShape',
			'smartArtTextPane',
			'smartArtPromote',
			'smartArtDemote',
			'smartArtRightToLeft',
			'smartArtMoveUp',
			'smartArtMoveDown',
			'smartArtResetGraphic',
			'smartArtConvert',
		] as const) {
			expect(buildRibbonGallery(id, { element: null }).disabled).toBeTruthy();
		}
	});

	it('add Shape appends a top-level node through the core edit', () => {
		const result = run('smartArtAddShape');
		expect(result?.kind).toBe('element');
		const data = (result as { patch: { smartArtData: { nodes: unknown[] } } }).patch.smartArtData;
		expect(data.nodes).toHaveLength(3);
	});

	it('add Shape Before inserts a top-level node before the first one', () => {
		const result = run('smartArtAddShapeBefore');
		const data = (result as { patch: { smartArtData: { nodes: { id: string }[] } } }).patch
			.smartArtData;
		expect(data.nodes).toHaveLength(3);
		expect(data.nodes[1]!.id).toBe('n1');
	});

	it('add Shape Above puts a new parent over the last top-level node', () => {
		const result = run('smartArtAddShapeAbove');
		const data = (
			result as { patch: { smartArtData: { nodes: { id: string; parentId?: string }[] } } }
		).patch.smartArtData;
		const added = data.nodes.find((node) => node.id !== 'n1' && node.id !== 'n2');
		expect(added).toBeDefined();
		expect(data.nodes.find((node) => node.id === 'n2')?.parentId).toBe(added!.id);
	});

	it('add Bullet adds a child under the last top-level node', () => {
		const result = run('smartArtAddBullet');
		const data = (result as { patch: { smartArtData: { nodes: { parentId?: string }[] } } }).patch
			.smartArtData;
		expect(data.nodes).toHaveLength(3);
		expect(data.nodes.some((node) => node.parentId === 'n2')).toBeTruthy();
	});

	it('text Pane emits the smartArtTextPane host command', () => {
		expect(run('smartArtTextPane')).toStrictEqual({
			kind: 'command',
			command: 'smartArtTextPane',
		});
	});

	it('promote elevates a child node and disables when no node can be promoted', () => {
		const flatDesc = buildRibbonGallery('smartArtPromote', ctx);
		expect(flatDesc.disabled).toBeTruthy();
		expect(flatDesc.command?.hintKey).toBe('pptx.gallery.smartArtCommand.needsNode');
		expect(run('smartArtPromote')).toBeNull();

		const result = run('smartArtPromote', { element: hierarchicalSmartArt });
		expect(result?.kind).toBe('element');
		const data = (result as { patch: { smartArtData: PptxSmartArtData } }).patch.smartArtData;
		expect(data.nodes.find((node) => node.id === 'n1a')?.parentId).toBeUndefined();
	});

	it('demote nests a node under its preceding sibling and respects selectedSmartArtNodeId', () => {
		expect(
			buildRibbonGallery('smartArtDemote', { element: singleNodeSmartArt }).disabled,
		).toBeTruthy();
		expect(
			buildRibbonGallery('smartArtDemote', { element: smartArt, selectedSmartArtNodeId: 'n1' })
				.disabled,
		).toBeTruthy();

		const result = run('smartArtDemote', { element: smartArt, selectedSmartArtNodeId: 'n2' });
		expect(result?.kind).toBe('element');
		const data = (result as { patch: { smartArtData: PptxSmartArtData } }).patch.smartArtData;
		expect(data.nodes.find((node) => node.id === 'n2')?.parentId).toBe('n1');
	});

	it('move Up and Move Down reorder siblings and respect boundaries', () => {
		expect(
			buildRibbonGallery('smartArtMoveUp', { element: singleNodeSmartArt }).disabled,
		).toBeTruthy();
		expect(
			buildRibbonGallery('smartArtMoveDown', { element: singleNodeSmartArt }).disabled,
		).toBeTruthy();
		expect(
			buildRibbonGallery('smartArtMoveUp', { element: smartArt, selectedSmartArtNodeId: 'n1' })
				.disabled,
		).toBeTruthy();
		expect(
			buildRibbonGallery('smartArtMoveDown', { element: smartArt, selectedSmartArtNodeId: 'n2' })
				.disabled,
		).toBeTruthy();

		const up = run('smartArtMoveUp');
		expect(up?.kind).toBe('element');
		const upNodes = (up as { patch: { smartArtData: PptxSmartArtData } }).patch.smartArtData.nodes;
		expect(upNodes.map((node) => node.id)).toStrictEqual(['n2', 'n1']);

		const down = run('smartArtMoveDown', { element: smartArt, selectedSmartArtNodeId: 'n1' });
		expect(down?.kind).toBe('element');
		const downNodes = (down as { patch: { smartArtData: PptxSmartArtData } }).patch.smartArtData
			.nodes;
		expect(downNodes.map((node) => node.id)).toStrictEqual(['n2', 'n1']);
	});

	it('right to Left toggles presLayoutVars.direction between rev and norm and rebuilds shapes', () => {
		const withShapes = {
			...smartArt,
			smartArtData: {
				...(smartArt as { smartArtData: PptxSmartArtData }).smartArtData,
				drawingShapes: [
					{ id: 's1', x: 0, y: 0, width: 100, height: 50, text: 'One' },
					{ id: 's2', x: 120, y: 0, width: 100, height: 50, text: 'Two' },
				],
			},
		} as unknown as PptxElement;
		const reversed = run('smartArtRightToLeft', { element: withShapes });
		expect(reversed?.kind).toBe('element');
		const revData = (reversed as { patch: { smartArtData: PptxSmartArtData } }).patch.smartArtData;
		expect(revData.presLayoutVars?.direction).toBe('rev');
		expect(revData.drawingShapes?.length).toBeGreaterThanOrEqual(2);
		expect(revData.drawingShapes?.[0]?.text).toBe('Two');

		const restored = run('smartArtRightToLeft', {
			element: { ...withShapes, smartArtData: revData } as PptxElement,
		});
		expect(restored?.kind).toBe('element');
		const normData = (restored as { patch: { smartArtData: PptxSmartArtData } }).patch.smartArtData;
		expect(normData.presLayoutVars?.direction).toBe('norm');
		expect(normData.drawingShapes?.[0]?.text).toBe('One');
	});

	it('convert decomposes the SmartArt into standard shapes via replaceElement', () => {
		const result = run('smartArtConvert');
		expect(result?.kind).toBe('replaceElement');
		const replaced = result as {
			kind: 'replaceElement';
			elementId: string;
			elements: PptxElement[];
		};
		expect(replaced.elementId).toBe('sa1');
		expect(replaced.elements.length).toBeGreaterThanOrEqual(2);
		expect(replaced.elements.every((el) => el.type !== 'smartArt')).toBeTruthy();
		expect(replaced.elements[0]!.x).toBeGreaterThanOrEqual(10);
		expect(replaced.elements[0]!.y).toBeGreaterThanOrEqual(20);
	});

	it('replaceElementInList splices replacements in place and preserves paint order', () => {
		const a = { id: 'a' } as PptxElement;
		const b = { id: 'b' } as PptxElement;
		const c = { id: 'c' } as PptxElement;
		const r1 = { id: 'r1' } as PptxElement;
		const r2 = { id: 'r2' } as PptxElement;
		expect(replaceElementInList([a, b, c], 'b', [r1, r2])?.map((el) => el.id)).toStrictEqual([
			'a',
			'r1',
			'r2',
			'c',
		]);
		expect(replaceElementInList([a, b], 'missing', [r1])).toBeNull();
	});
});

describe('smartArt Layouts gallery', () => {
	it('lists the switchable families and marks the current one', () => {
		const descriptor = buildRibbonGallery('smartArtLayouts', ctx);
		const items = descriptor.sections[0].items;
		expect(items.length).toBeGreaterThan(5);
		expect(items.filter((item) => item.applied).map((item) => item.id)).toStrictEqual(['list']);
		expect(items.every((item) => item.previewSvg.startsWith('<svg'))).toBeTruthy();
	});

	it('switches layout like the Inspector, clearing stale drawing shapes', () => {
		const result = applyRibbonGalleryItem('smartArtLayouts', 'cycle', ctx);
		expect(result?.kind).toBe('element');
		const data = (
			result as { patch: { smartArtData: { resolvedLayoutType?: string; nodes: unknown[] } } }
		).patch.smartArtData;
		expect(data.resolvedLayoutType).toBe('cycle');
		expect(data.nodes).toHaveLength(2);
		expect(applyRibbonGalleryItem('smartArtLayouts', 'list', ctx)).toBeNull();
	});
});
