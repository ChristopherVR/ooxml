import {
	getVisioPageLayers,
	type VisioDocument,
	type VisioPage,
	type VisioShape,
	type VisioMatrix,
} from 'ooxml-core/visio';
import { fillPaint } from './render-fill.js';
import { renderImage } from './render-image.js';
import { RenderResources } from './render-resources.js';
import { applyArrowheads } from './arrowheads.js';
import { assertViewableDocument } from './scene-validation.js';
import { createTextLayoutBudget, type TextLayoutBudget } from './text-layout.js';
import { renderText } from './render-text.js';

const NS = 'http://www.w3.org/2000/svg';
export function svgElement<K extends keyof SVGElementTagNameMap>(name: K): SVGElementTagNameMap[K] {
	return document.createElementNS(NS, name);
}
export function safeColor(value: string, fallback = '#273c46'): string {
	return /^(#[\da-f]{3,8}|[a-z]+|(?:rgb|hsl)a?\([\d.,%\s+-]+\))$/i.test(value) ? value : fallback;
}
export function matrix(value: VisioMatrix): string {
	return `matrix(${value.map((number) => (Number.isFinite(number) ? number : 0)).join(' ')})`;
}
export interface RenderResult {
	svg: SVGSVGElement;
	warnings: string[];
	dispose(): void;
}
export interface RenderOptions {
	/** Static output embeds rasters and omits viewer selection semantics. Prefer exportPageSvg for bounded serialization. */
	static?: boolean;
}
interface RenderContext {
	defs: SVGDefsElement;
	warnings: Set<string>;
	resources: RenderResources;
	nodes: number;
	textBudget: TextLayoutBudget;
	renderable: WeakMap<VisioShape, boolean>;
	interactive: boolean;
}
export function renderPage(
	model: VisioDocument,
	page: VisioPage,
	options: RenderOptions = {},
): RenderResult {
	assertViewableDocument(model);
	if (!model.pages.includes(page))
		throw new Error('The selected page does not belong to this document.');
	const svg = svgElement('svg');
	if (!options.static) svg.classList.add('paper');
	svg.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns', NS);
	if (options.static)
		svg.setAttributeNS(
			'http://www.w3.org/2000/xmlns/',
			'xmlns:xlink',
			'http://www.w3.org/1999/xlink',
		);
	svg.setAttribute('viewBox', `0 0 ${page.width} ${page.height}`);
	svg.setAttribute('role', options.static ? 'img' : 'group');
	svg.setAttribute('aria-label', page.name);
	const title = svgElement('title');
	title.textContent = page.name;
	svg.append(title);
	const defs = svgElement('defs');
	svg.append(defs);
	const root = svgElement('g');
	root.setAttribute('transform', `translate(0 ${page.height}) scale(1 -1)`);
	svg.append(root);
	const context: RenderContext = {
		defs,
		warnings: new Set(),
		resources: new RenderResources(options.static ? defs : undefined),
		nodes: 0,
		textBudget: createTextLayoutBudget(),
		renderable: new WeakMap(),
		interactive: !options.static,
	};
	try {
		for (const layer of getVisioPageLayers(model, page.id))
			for (const shape of layer.shapes) drawShape(shape, root, context, layer.id);
	} catch (error) {
		context.resources.dispose();
		throw error;
	}
	if (!options.static) {
		const first = svg.querySelector<SVGGElement>('[data-shape-id]');
		first?.setAttribute('tabindex', '0');
	}
	return { svg, warnings: [...context.warnings], dispose: () => context.resources.dispose() };
}
function hasRenderableContent(shape: VisioShape, context: RenderContext): boolean {
	const cached = context.renderable.get(shape);
	if (cached !== undefined) return cached;
	const own =
		!(shape.kind === 'group' && shape.groupDisplayMode === 0) &&
		(shape.geometry.length > 0 ||
			!!shape.image ||
			!!shape.text.plainText ||
			shape.kind === 'foreign');
	const visible =
		!shape.hidden && (own || shape.children.some((child) => hasRenderableContent(child, context)));
	context.renderable.set(shape, visible);
	return visible;
}
function drawShape(
	shape: VisioShape,
	parent: SVGElement,
	context: RenderContext,
	pageId: string,
): void {
	if (!hasRenderableContent(shape, context)) return;
	if (++context.nodes > 50_000) {
		context.warnings.add(
			'Some drawing content was omitted because it exceeds the safe rendering node limit.',
		);
		return;
	}
	const group = svgElement('g');
	group.setAttribute('transform', matrix(shape.transform));
	if (context.interactive) {
		group.dataset.shapeId = shape.id;
		group.dataset.shapeName = shape.name;
		group.dataset.pageId = pageId;
		group.setAttribute('role', shape.kind === 'group' ? 'group' : 'button');
		group.setAttribute('tabindex', '-1');
		group.setAttribute('aria-label', shape.text.plainText || shape.name || `Shape ${shape.id}`);
	}
	const title = svgElement('title');
	title.textContent = shape.text.plainText || shape.name;
	group.append(title);
	parent.append(group);
	const own = () => drawOwn(shape, group, context);
	const children = () => {
		for (const child of shape.children) drawShape(child, group, context, pageId);
	};
	if (shape.kind === 'group' && shape.groupDisplayMode === 0) {
		children();
		return;
	}
	if (shape.kind === 'group' && shape.groupDisplayMode === 2) {
		children();
		own();
		return;
	}
	own();
	children();
}
function drawOwn(shape: VisioShape, group: SVGElement, context: RenderContext): void {
	const { warnings, defs, resources } = context;
	const fill = shape.geometry.some((geometry) => geometry.fill)
		? fillPaint(shape.style, defs)
		: 'none';
	for (const geometry of shape.geometry) {
		if (++context.nodes > 50_000) {
			warnings.add(
				'Some drawing content was omitted because it exceeds the safe rendering node limit.',
			);
			return;
		}
		const path = svgElement('path');
		if (context.interactive) path.dataset.geometry = '';
		path.setAttribute('d', geometry.path);
		path.setAttribute('fill', geometry.fill ? fill : 'none');
		path.setAttribute('fill-opacity', String(shape.style.fillOpacity));
		path.setAttribute(
			'stroke',
			geometry.stroke && shape.style.linePattern !== 0 ? safeColor(shape.style.lineColor) : 'none',
		);
		path.setAttribute('stroke-width', String(shape.style.lineWidth));
		path.setAttribute('stroke-opacity', String(shape.style.lineOpacity));
		path.setAttribute('stroke-linejoin', 'round');
		path.setAttribute('stroke-linecap', shape.style.lineCap ?? 'round');
		if (shape.style.linePattern > 1) {
			const unit = Math.max(shape.style.lineWidth, 0.01);
			const patterns: Record<number, number[]> = {
				2: [6, 3],
				3: [1, 3],
				4: [6, 3, 1, 3],
				5: [6, 3, 1, 3, 1, 3],
			};
			const dash = patterns[shape.style.linePattern];
			if (dash) path.setAttribute('stroke-dasharray', dash.map((n) => n * unit).join(' '));
			warnings.add('Visio line-pattern spacing is approximated.');
		}
		applyArrowheads(path, shape.style, defs, warnings);
		group.append(path);
	}
	const raster = renderImage(shape, resources);
	if (raster) group.append(raster);
	if (shape.text.plainText) group.append(renderText(shape.text, warnings, context.textBudget));
	if (shape.kind === 'foreign' && !shape.image)
		warnings.add('Embedded or foreign objects are not rendered in this build.');
}
