import { afterEach, describe, expect, it, vi } from 'vitest';
import { sanitizeVisioForeignVectorTree, type VisioDocument } from 'ooxml-core/visio';
import { demoDocument } from './demo-document.js';
import { renderPage } from './render-svg.js';
import { exportPageSvg } from './export-svg.js';
import { createPrintSnapshot } from './print-snapshot.js';
import { copySnapshotScene } from './snapshot-scene.js';
import { assertViewableDocument } from './scene-validation.js';
import { inspectForeignVectorResource } from './foreign-vector-budget.js';
import { mountViewer } from './binding.js';

function vector() {
	return sanitizeVisioForeignVectorTree({
		tag: 'svg',
		attrs: { xmlns: 'http://www.w3.org/2000/svg', width: 30, height: 20, viewBox: '0 0 30 20' },
		children: [
			{
				tag: 'defs',
				attrs: {},
				children: [
					{
						tag: 'clipPath',
						attrs: { id: 'left' },
						children: [{ tag: 'path', attrs: { d: 'M0 0 L15 0 L15 20 L0 20 Z' } }],
					},
				],
			},
			{
				tag: 'g',
				attrs: { 'clip-path': 'url(#left)' },
				children: [
					{ tag: 'path', attrs: { d: 'M0 0 L30 0 L0 20 Z', fill: '#ff0000', stroke: 'none' } },
				],
			},
		],
	});
}
function scene(): VisioDocument {
	const model = structuredClone(demoDocument),
		page = model.pages[0]!,
		shape = page.shapes[0]!;
	model.pages = [page];
	page.shapes = [shape];
	page.width = page.height = 4;
	shape.kind = 'foreign';
	shape.geometry = [];
	shape.text.plainText = '';
	shape.text.runs = [];
	shape.width = 2;
	shape.height = 1;
	shape.foreignVector = { vector: vector(), x: -0.25, y: 0.25, width: 3, height: 2, opacity: 0.6 };
	return model;
}
const parse = (svg: string) =>
	new DOMParser().parseFromString(svg, 'image/svg+xml').documentElement;
const output = {
	live: (model: VisioDocument) => renderPage(model, model.pages[0]!).svg,
	export: (model: VisioDocument) => parse(exportPageSvg(model).svg),
	print: (model: VisioDocument) =>
		parse(createPrintSnapshot(model, { pageIndices: [0] }).pages[0]!.svg),
};
afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

describe('placed foreign vectors through the shared scene renderer', () => {
	it.each(Object.entries(output))(
		'%s preserves placement, orientation, opacity and both clips',
		(_name, render) => {
			const model = scene(),
				shape = model.pages[0]!.shapes[0]!,
				svg = render(model);
			const nested = svg.querySelector('svg')!;
			expect(nested.getAttribute('viewBox')).toBe('0 0 30 20');
			expect(nested.getAttribute('width')).toBe('3');
			expect(nested.getAttribute('height')).toBe('2');
			expect(nested.getAttribute('transform')).toBe('translate(-0.25 2.25) scale(1 -1)');
			expect(nested.getAttribute('opacity')).toBe('0.6');
			expect(nested.getAttribute('overflow')).toBe('hidden');
			expect(nested.getAttribute('preserveAspectRatio')).toBe('none');
			const frame = svg.querySelector('clipPath rect')!;
			expect(frame.getAttribute('width')).toBe(String(shape.width));
			expect(frame.getAttribute('height')).toBe(String(shape.height));
			expect(nested.parentElement!.getAttribute('transform')).toBeNull();
			expect(svg.querySelectorAll('clipPath')).toHaveLength(2);
			const ids = new Set([...svg.querySelectorAll('[id]')].map((node) => node.id));
			for (const node of svg.querySelectorAll('[clip-path]'))
				expect(ids.has(node.getAttribute('clip-path')!.slice(5, -1))).toBe(true);
			expect(ids.has('left')).toBe(false);
			expect(svg.querySelector('script,foreignObject,image,use,a,style,animate')).toBeNull();
			expect(svg.textContent).not.toContain('foreign objects are not rendered');
		},
	);
	it('uses unique closed clip references for shared vector instances and leaves the host unchanged', () => {
		const model = scene(),
			shape = model.pages[0]!.shapes[0]!;
		model.pages[0]!.shapes.push({ ...shape, id: 'again' });
		const before = structuredClone(model),
			{ svg, warnings } = renderPage(model, model.pages[0]!);
		expect(svg.querySelectorAll('svg')).toHaveLength(2);
		const ids = [...svg.querySelectorAll('[id]')].map((node) => node.id);
		expect(new Set(ids).size).toBe(ids.length);
		expect(warnings).toEqual([]);
		expect(model).toEqual(before);
	});
	it('keeps vector-only host shapes selectable and shares live/export/print output', () => {
		const model = scene(),
			shape = model.pages[0]!.shapes[0]!;
		shape.kind = 'shape';
		const host = document.createElement('div');
		document.body.append(host);
		const viewer = mountViewer(host, { document: model });
		expect(viewer.element.shadowRoot!.querySelector('[data-shape-id] svg')).not.toBeNull();
		viewer.controller.selectShape({ id: shape.id, name: shape.name });
		const before = viewer.controller.state;
		expect(parse(viewer.exportSvg().svg).querySelector('svg')).not.toBeNull();
		expect(parse(viewer.createPrintSnapshot().pages[0]!.svg).querySelector('svg')).not.toBeNull();
		expect(viewer.controller.state).toBe(before);
		viewer.destroy();
	});
	it.each(Object.entries(output))(
		'%s honors hidden and group-content suppression',
		(_name, render) => {
			const model = scene(),
				shape = model.pages[0]!.shapes[0]!;
			shape.hidden = true;
			expect(render(model).querySelector('svg')).toBeNull();
			shape.hidden = false;
			shape.kind = 'group';
			shape.groupDisplayMode = 0;
			expect(render(model).querySelector('svg')).toBeNull();
			shape.groupDisplayMode = 1;
			expect(render(model).querySelector('svg')).not.toBeNull();
		},
	);
	it('temporary layer overrides affect live vectors while snapshots preserve saved display', () => {
		const model = scene(),
			page = model.pages[0]!,
			shape = page.shapes[0]!;
		page.layers = [{ id: 'layer', name: 'Hidden', visible: false, printable: true, locked: false }];
		shape.layerIds = ['layer'];
		shape.hidden = true;
		shape.visibility = {
			layerHidden: true,
			guide: false,
			noShow: false,
			layerPrintSummary: 'all-enabled',
		};
		const rendered = renderPage(model, page, {
			layerVisibilityOverrides: [{ pageId: page.id, layerId: 'layer', visible: true }],
		});
		expect(rendered.svg.querySelector('svg')).not.toBeNull();
		expect(output.export(model).querySelector('svg')).toBeNull();
		expect(output.print(model).querySelector('svg')).toBeNull();
	});
});

describe('untrusted vector scenes and aggregate work', () => {
	it.each(Object.entries(output))(
		'%s rejects a frozen forged vector before DOM allocation',
		(_name, render) => {
			const model = scene(),
				placed = model.pages[0]!.shapes[0]!.foreignVector!;
			const forged = structuredClone(placed.vector);
			Object.assign(forged.items[0]!, { clipIndex: 999 });
			placed.vector = Object.freeze(forged);
			const allocate = vi.spyOn(document, 'createElementNS');
			expect(() => render(model)).toThrow('clip index');
			expect(allocate).not.toHaveBeenCalled();
		},
	);
	it('rejects active paints, vector accessors, cycles and malformed placement', () => {
		const model = scene(),
			placed = model.pages[0]!.shapes[0]!.foreignVector!;
		const forged = structuredClone(placed.vector),
			group = forged.items[0]!;
		if (group.kind !== 'group' || group.items[0]!.kind !== 'path') throw new Error('fixture');
		Object.assign(group.items[0]!.paint, { fill: 'url(https://invalid.example/p)' });
		placed.vector = forged;
		expect(() => assertViewableDocument(model)).toThrow();
		const getter = vi.fn(() => []);
		Object.defineProperty(forged, 'items', { enumerable: true, get: getter });
		expect(() => assertViewableDocument(model)).toThrow();
		expect(getter).not.toHaveBeenCalled();
		placed.vector = structuredClone(vector());
		Object.assign(placed.vector.clips[0]!, { clipIndex: 0 });
		expect(() => assertViewableDocument(model)).toThrow('Cyclic');
		placed.vector = vector();
		for (const [key, value] of [
			['x', Infinity],
			['y', '0'],
			['width', -1],
			['height', NaN],
			['opacity', 2],
		]) {
			const old = Object.getOwnPropertyDescriptor(placed, key!)!;
			Object.assign(placed, { [key as string]: value });
			expect(() => assertViewableDocument(model)).toThrow('foreign vector');
			Object.defineProperty(placed, key!, old);
		}
	});
	it('charges repeated clip expansion for each instance including hidden shapes and snapshot copying', () => {
		const model = scene(),
			shape = model.pages[0]!.shapes[0]!;
		const path = 'M0 0 ' + 'L1 1 '.repeat(249);
		shape.foreignVector!.vector = sanitizeVisioForeignVectorTree({
			tag: 'svg',
			attrs: { xmlns: 'http://www.w3.org/2000/svg', width: 30, height: 20, viewBox: '0 0 30 20' },
			children: [
				{
					tag: 'defs',
					attrs: {},
					children: [
						{
							tag: 'clipPath',
							attrs: { id: 'cost' },
							children: [{ tag: 'path', attrs: { d: path } }],
						},
					],
				},
				...Array.from({ length: 100 }, () => ({
					tag: 'path',
					attrs: { d: 'M0 0', 'clip-path': 'url(#cost)' },
				})),
			],
		});
		const cost = inspectForeignVectorResource(shape.foreignVector!.vector);
		expect(cost.operands).toBe(50_700);
		shape.hidden = true;
		model.pages[0]!.shapes = Array.from({ length: 20 }, (_, index) => ({
			...shape,
			id: String(index),
		}));
		const allocate = vi.spyOn(document, 'createElementNS');
		expect(() => assertViewableDocument(model)).toThrow('aggregate foreign vector');
		expect(() => copySnapshotScene(model)).toThrow('aggregate foreign vector');
		expect(() => output.live(model)).toThrow('aggregate foreign vector');
		expect(allocate).not.toHaveBeenCalled();
	});
	it('preflights vector serialization before allocating SVG nodes', () => {
		const model = scene(),
			allocate = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model, 0, { maxBytes: 10_000 })).toThrow('amplification');
		expect(allocate).not.toHaveBeenCalled();
	});
	it('revalidates the same mutable vector on subsequent calls', () => {
		const model = scene(),
			placed = model.pages[0]!.shapes[0]!.foreignVector!;
		placed.vector = structuredClone(placed.vector);
		assertViewableDocument(model);
		Object.assign(placed.vector, { width: NaN });
		expect(() => assertViewableDocument(model)).toThrow();
	});
});

describe('vector print capture and whole-job budgets', () => {
	it('detaches all vector state and does not freeze host-owned objects', () => {
		const model = scene(),
			placed = model.pages[0]!.shapes[0]!.foreignVector!;
		placed.vector = structuredClone(placed.vector);
		const copied = copySnapshotScene(model).pages[0]!.shapes[0]!.foreignVector!;
		expect(copied).toEqual(placed);
		expect(copied).not.toBe(placed);
		expect(copied.vector).not.toBe(placed.vector);
		expect(copied.vector.items).not.toBe(placed.vector.items);
		expect(copied.vector.clips[0]!.items[0]!.commands).not.toBe(
			placed.vector.clips[0]!.items[0]!.commands,
		);
		expect(Object.isFrozen(placed.vector)).toBe(false);
		const create = document.createElementNS.bind(document);
		vi.spyOn(document, 'createElementNS').mockImplementationOnce((namespace, name, options) => {
			placed.x = 999;
			placed.width = 0;
			Object.assign(placed.vector, { width: NaN, items: [] });
			return create(namespace, name, options);
		});
		const svg = output.print(model),
			nested = svg.querySelector('svg')!;
		expect(nested.getAttribute('width')).toBe('3');
		expect(nested.getAttribute('transform')).toBe('translate(-0.25 2.25) scale(1 -1)');
		expect(nested.querySelector('path[fill]')!.getAttribute('fill')).toBe('#ff0000');
	});
	it('accounts repeated background instances and accepts exact lowered vector budgets', () => {
		const model = scene(),
			background = model.pages[0]!;
		background.isBackground = true;
		model.pages.push({
			...background,
			id: 'front-a',
			isBackground: false,
			backgroundPageId: background.id,
			shapes: [],
		});
		model.pages.push({ ...model.pages[1]!, id: 'front-b' });
		const cost = inspectForeignVectorResource(background.shapes[0]!.foreignVector!.vector);
		const result = createPrintSnapshot(model, { pageIndices: [1, 2] });
		expect(result.usage.vectorNodes).toBe(cost.nodes * 2);
		expect(result.usage.vectorCommands).toBe(cost.commands * 2);
		expect(result.usage.vectorOperands).toBe(cost.operands * 2);
		const limits = {
			maxVectorNodes: cost.nodes * 2,
			maxVectorCommands: cost.commands * 2,
			maxVectorOperands: cost.operands * 2,
		};
		expect(createPrintSnapshot(model, { pageIndices: [1, 2], limits }).pages).toHaveLength(2);
		const allocate = vi.spyOn(document, 'createElementNS');
		for (const key of Object.keys(limits) as (keyof typeof limits)[])
			expect(() =>
				createPrintSnapshot(model, { pageIndices: [1, 2], limits: { [key]: limits[key] - 1 } }),
			).toThrow(/vector/);
		expect(allocate).not.toHaveBeenCalled();
	});
	it('includes vectors on unselected pages in repeated validation work', () => {
		const model = scene(),
			vectorPage = model.pages[0]!;
		model.pages.push({ ...vectorPage, id: 'blank', shapes: [] });
		const baseline = createPrintSnapshot(
			{ ...model, pages: [model.pages[1]!] },
			{ pageIndices: [0] },
		);
		const result = createPrintSnapshot(model, { pageIndices: [1] });
		expect(result.usage.vectorNodes).toBe(0);
		expect(result.usage.validationWork).toBeGreaterThan(baseline.usage.validationWork);
		const allocate = vi.spyOn(document, 'createElementNS');
		expect(() =>
			createPrintSnapshot(model, {
				pageIndices: [1],
				limits: { maxValidationWork: baseline.usage.validationWork },
			}),
		).toThrow('validation work');
		expect(allocate).not.toHaveBeenCalled();
	});
});
