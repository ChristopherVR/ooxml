import { describe, expect, it } from 'vitest';
import { sanitizeVisioForeignVectorTree } from 'ooxml-core/visio';
import { renderForeignVector } from './render-foreign-vector.js';
const source = () => ({
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
					children: [{ tag: 'path', attrs: { d: 'M0 0 L10 0 L10 20 L0 20 Z' } }],
				},
				{
					tag: 'clipPath',
					attrs: { id: 'top', 'clip-path': 'url(#left)' },
					children: [{ tag: 'path', attrs: { d: 'M0 0 L30 0 L30 10 L0 10 Z' } }],
				},
			],
		},
		{
			tag: 'g',
			attrs: { 'clip-path': 'url(#top)' },
			children: [
				{ tag: 'path', attrs: { d: 'M0 0 L30 0 L30 20 L0 20 Z', fill: '#ff0000', stroke: 'none' } },
			],
		},
	],
});
describe('inert foreign vector renderer boundary', () => {
	it('builds fixed SVG nodes with closed generated clip identities', () => {
		const svg = renderForeignVector(sanitizeVisioForeignVectorTree(source()));
		expect(svg.tagName).toBe('svg');
		expect(svg.getAttribute('viewBox')).toBe('0 0 30 20');
		expect(svg.getAttribute('overflow')).toBe('hidden');
		expect(svg.querySelectorAll('clipPath')).toHaveLength(2);
		expect(svg.querySelectorAll('path')).toHaveLength(3);
		const ids = new Set([...svg.querySelectorAll('[id]')].map((x) => x.id));
		expect(ids.has('left')).toBe(false);
		expect(ids.has('top')).toBe(false);
		for (const node of svg.querySelectorAll('[clip-path]'))
			expect(ids.has(node.getAttribute('clip-path')!.slice(5, -1))).toBe(true);
		expect(svg.querySelector('g path')?.getAttribute('fill')).toBe('#ff0000');
	});
	it('assigns distinct DOM resource identities to separate identical scenes', () => {
		const scene = sanitizeVisioForeignVectorTree(source());
		const a = renderForeignVector(scene),
			b = renderForeignVector(scene);
		const left = new Set([...a.querySelectorAll('[id]')].map((x) => x.id));
		expect([...b.querySelectorAll('[id]')].some((x) => left.has(x.id))).toBe(false);
	});
	it('revalidates forged canonical payloads instead of trusting TypeScript or freezing', () => {
		const scene = sanitizeVisioForeignVectorTree(source());
		const forged = JSON.parse(JSON.stringify(scene));
		forged.items[0].items[0].paint.fill = 'url(https://example.invalid/a)';
		expect(() => renderForeignVector(Object.freeze(forged))).toThrow();
		const dangling = JSON.parse(JSON.stringify(scene));
		dangling.items[0].clipIndex = 999;
		expect(() => renderForeignVector(dangling)).toThrow();
	});
	it('produces no external references, event attributes, HTML, scripts or source IDs', () => {
		const svg = renderForeignVector(sanitizeVisioForeignVectorTree(source()));
		const serialized = new XMLSerializer().serializeToString(svg);
		const parsed = new DOMParser().parseFromString(serialized, 'image/svg+xml');
		expect(parsed.querySelector('parsererror')).toBeNull();
		expect(svg.querySelector('script,foreignObject,image,use,a,text,style,animate')).toBeNull();
		for (const node of svg.querySelectorAll('*'))
			for (const attr of node.getAttributeNames())
				expect(attr.toLowerCase().startsWith('on')).toBe(false);
		expect(serialized).not.toContain('https://example');
	});
});
