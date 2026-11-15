// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';

import { prepareExportClone } from '../render/export-clone';
import { inlineComputedStylesOnClone } from '../render/foreign-object-style-inline';
import { isolateLive } from './pdf-vector-raster';
import { prepareIsolatedRasterBase, RASTER_TARGET_ATTRIBUTE } from './pdf-vector-raster-batch';
import { isolatedMarkup } from './pdf-vector-raster-isolate';
import { RASTER_VISIBILITY_ATTRIBUTE, shadowTargetStyle } from './pdf-vector-raster-live';

const SVG = 'http://www.w3.org/2000/svg';

/**
 * A stand-in for `getComputedStyle`: the element's `data-css` as its computed style, so a test
 * sets "stylesheet" values apart from what its `style` attribute declares.
 */
const computed = (el: Element) => {
	const style = document.createElement('div').style;
	style.cssText = el.getAttribute('data-css') ?? '';
	return style;
};

function add(parent: Element, css: string, text?: string): HTMLElement {
	const el = document.createElement('div');
	el.setAttribute('data-css', css);
	if (text) el.textContent = text;
	parent.append(el);
	return el;
}

function stage(): HTMLElement {
	const root = add(document.body, 'position: relative; visibility: visible; overflow: hidden');
	return root;
}

const shape = (root: Element, css = '') =>
	add(root, `position: absolute; visibility: visible; ${css}`);

/** The markup with every visibility slot removed, to compare with a plain serialisation. */
const withoutSlots = (markup: string) =>
	markup.replaceAll(/;?visibility:(?:visible|hidden)!important/gu, '');

afterEach(() => {
	document.body.replaceChildren();
});

describe('prepareIsolatedRasterBase', () => {
	it('finds each target in the clone and untags the live stage', async () => {
		const root = stage();
		const a = shape(root);
		const b = shape(root);
		b.append(document.createElement('span'));
		const targets = new Map([
			[a, '0'],
			[b, '1'],
		]);
		const base = (await prepareIsolatedRasterBase(root, targets, { readComputedStyle: computed }))!;
		expect(base.targets.get('0')).toBe(1);
		expect(base.targets.get('1')).toBe(2);
		expect(
			root.querySelector(`[${RASTER_TARGET_ATTRIBUTE}], [${RASTER_VISIBILITY_ATTRIBUTE}]`),
		).toBeNull();
		expect(root.hasAttribute(RASTER_VISIBILITY_ATTRIBUTE)).toBe(false);
		const markup = isolatedMarkup(base.template, 1, 'subtree', { prune: false });
		expect(markup).not.toContain(RASTER_TARGET_ATTRIBUTE);
		expect(markup).not.toContain(RASTER_VISIBILITY_ATTRIBUTE);
		expect(markup).not.toContain('<!--');
		expect(markup).not.toMatch(/--p[a-z0-9]+-\d/u);
	});

	it('serialises the clone exactly, apart from the visibility it writes', async () => {
		const root = stage();
		const a = shape(root);
		a.textContent = 'Tom & "Jerry" <3';
		add(root, 'position: static; visibility: visible').append(document.createElement('br'));
		const base = (await prepareIsolatedRasterBase(root, new Map([[a, '0']]), {
			readComputedStyle: computed,
		}))!;
		const plain = root.cloneNode(true) as HTMLElement;
		inlineComputedStylesOnClone(root, plain, computed);
		prepareExportClone(plain);
		// The stage is drawn without its own transform (see the zoom test below).
		for (const property of ['transform', 'translate', 'rotate', 'scale'])
			plain.style.setProperty(property, 'none');
		const expected = new XMLSerializer().serializeToString(plain);
		expect(withoutSlots(isolatedMarkup(base.template, 1, 'subtree', { prune: false }))).toBe(
			expected,
		);
	});

	it('leaves out only hidden, out-of-flow subtrees nothing else depends on', async () => {
		const root = stage();
		const target = shape(root);
		shape(root).append('other shape');
		add(root, 'position: static; visibility: visible', 'in flow');
		shape(root, 'counter-reset: n').append('counted');
		const svg = document.createElementNS(SVG, 'svg');
		const gradient = document.createElementNS(SVG, 'linearGradient');
		gradient.setAttribute('id', 'grad-1');
		svg.append(gradient);
		shape(root).append(svg);
		shape(root, 'background-image: url(#grad-1)');
		const scroller = shape(root, 'overflow: auto');
		add(scroller, 'position: absolute; visibility: visible', 'scrolled');
		const base = (await prepareIsolatedRasterBase(root, new Map([[target, '0']]), {
			readComputedStyle: computed,
		}))!;
		const markup = isolatedMarkup(base.template, 1, 'subtree');
		expect(markup).not.toContain('other shape');
		expect(markup).toContain('in flow');
		expect(markup).toContain('counted');
		expect(markup).toContain('grad-1');
		// The scroll container itself is out of flow and may go; what it scrolls may not on its own.
		expect(markup).not.toContain('scrolled');
		expect(isolatedMarkup(base.template, 1, 'subtree', { prune: false })).toContain('other shape');
		const last = base.template.parents.length - 1;
		expect(isolatedMarkup(base.template, last, 'subtree')).toContain('scrolled');
	});

	it('draws only the target: a shape declaring itself visible stays out', async () => {
		const root = stage();
		const target = shape(root);
		const inside = add(target, 'position: absolute; visibility: hidden');
		inside.style.visibility = 'visible';
		inside.append('declared inside');
		const declared = shape(root);
		declared.style.visibility = 'visible';
		declared.append('declared outside');
		const base = (await prepareIsolatedRasterBase(root, new Map([[target, '0']]), {
			readComputedStyle: computed,
		}))!;
		for (const mode of ['subtree', 'self', 'shadow'] as const) {
			const markup = isolatedMarkup(base.template, 1, mode);
			expect(markup, mode).not.toContain('declared outside');
			// Inside a subtree picture what an element declares holds; self and shadow hide it.
			if (mode === 'subtree')
				expect(markup).toMatch(/visibility:visible!important">declared inside/u);
			else expect(markup, mode).not.toMatch(/visibility:visible!important">declared inside/u);
		}
		expect(root.querySelector(`[${RASTER_VISIBILITY_ATTRIBUTE}]`)).toBeNull();
		expect(declared.getAttribute('style')).toBe('visibility: visible;');
	});

	it('draws a zoomed stage at its own layout size, where the picture regions are measured', async () => {
		const root = stage();
		root.style.transform = 'scale(0.871875)';
		root.style.transformOrigin = '0px 0px';
		const target = shape(root);
		// A plain declaration: jsdom's own does not keep every transform value.
		const zoom = (el: Element) => {
			const base = computed(el);
			const map = new Map<string, string>();
			for (let i = 0; i < base.length; i++)
				map.set(base.item(i), base.getPropertyValue(base.item(i)));
			if (el === root) map.set('transform', 'matrix(0.871875, 0, 0, 0.871875, 0, 0)');
			if (el === target) map.set('transform', 'matrix(0.98, 0.17, -0.17, 0.98, 0, 0)');
			const keys = [...map.keys()];
			return {
				length: keys.length,
				item: (i: number) => keys[i] ?? '',
				getPropertyValue: (property: string) => map.get(property) ?? '',
			};
		};
		const base = (await prepareIsolatedRasterBase(root, new Map([[target, '0']]), {
			readComputedStyle: zoom,
		}))!;
		const markup = isolatedMarkup(base.template, 1, 'self');
		const [rootStyle = '', targetStyle = ''] = [...markup.matchAll(/ style="([^"]*)"/gu)].map(
			(match) => match[1],
		);
		expect(rootStyle).toMatch(/transform: none/u);
		expect(rootStyle).not.toContain('0.871875');
		// The content keeps its own transforms.
		expect(targetStyle).toContain('matrix(0.98, 0.17, -0.17, 0.98, 0, 0)');
		expect(root.style.transform).toBe('scale(0.871875)');
	});
});

describe('isolateLive', () => {
	it('hides shapes that declare themselves visible, and puts every style back', () => {
		const root = stage();
		const target = shape(root);
		const child = add(target, '');
		child.style.visibility = 'visible';
		const other = shape(root);
		other.style.visibility = 'visible';
		const restore = isolateLive(root, target, 'subtree');
		expect(root.style.visibility).toBe('hidden');
		expect(target.style.visibility).toBe('visible');
		expect(child.style.visibility).toBe('visible');
		expect(other.style.visibility).toBe('hidden');
		restore();
		expect(other.style.visibility).toBe('visible');
		expect(other.style.getPropertyPriority('visibility')).toBe('');
		expect(root.style.visibility).toBe('');
		const restoreSelf = isolateLive(root, target, 'self');
		expect(child.style.visibility).toBe('hidden');
		restoreSelf();
		expect(child.style.visibility).toBe('visible');
	});
});

describe('shadowTargetStyle', () => {
	it('reads the target as the shadow capture isolates it, then restores the live stage', () => {
		const root = stage();
		const target = shape(root);
		target.setAttribute('style', 'color: red; font-family: "Segoe UI"');
		const before = target.style.cssText;
		const style = shadowTargetStyle(root, target);
		expect(style).toContain('!important');
		expect(style).toMatch(/;color:(?:transparent|rgba\(0, 0, 0, 0\))!important/u);
		expect(style).not.toContain('"');
		expect(target.style.cssText).toBe(before);
		expect(root.style.visibility).toBe('');
	});
});
