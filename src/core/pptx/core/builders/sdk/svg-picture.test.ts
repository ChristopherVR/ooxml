/**
 * An SVG picture is saved the Office way, a raster blip with the SVG in an
 * `asvg:svgBlip` extension, and loads back as an SVG picture.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { ImagePptxElement } from '../../types';
import { createImageElement } from './ElementFactory';
import type { ImageOptions } from './types';

const SVG_TEXT =
	'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="red"/></svg>';
const SVG = `data:image/svg+xml;base64,${Buffer.from(SVG_TEXT).toString('base64')}`;
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function save(sources: Array<[string, ImageOptions?]>) {
	const { handler, createSlide } = await PptxHandler.create();
	const slide = createSlide('Blank');
	for (const [source, options] of sources) {
		slide.addImage(source, options);
	}
	const bytes = await handler.save([slide.build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const rels = (await zip.file('ppt/slides/_rels/slide1.xml.rels')?.async('string')) ?? '';
	const target = (id: string) =>
		new RegExp(`Id="${id}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${id}"`, 'u')
			.exec(rels)
			?.slice(1)
			.find(Boolean);
	const loaded = await new PptxHandler().load(bytes);
	return { zip, xml, target, pictures: (loaded.slides[0]?.elements ?? []) as ImagePptxElement[] };
}

const blips = (xml: string) => xml.match(/<a:blip\b[\s\S]*?<\/a:blip>/gu) ?? [];
const svgEmbed = (blip: string) => /<asvg:svgBlip [^>]*r:embed="([^"]+)"/u.exec(blip)?.[1] ?? '';

describe('SVG pictures', () => {
	it('embeds a PNG in the blip and the SVG in asvg:svgBlip', async () => {
		const { zip, xml, target, pictures } = await save([[SVG]]);
		const [blip = ''] = blips(xml);
		const raster = /<a:blip r:embed="([^"]+)"/u.exec(blip)?.[1] ?? '';
		expect(blip).toContain('uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"');
		expect(blip).toContain(
			'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"',
		);
		expect(target(raster)).toMatch(/\.png$/u);
		const svgTarget = target(svgEmbed(blip)) ?? '';
		expect(svgTarget).toMatch(/\.svg$/u);
		const svgPart = await zip.file(`ppt/${svgTarget.replace(/^\.\.\//u, '')}`)?.async('string');
		expect(svgPart).toBe(SVG_TEXT);
		const types = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
		expect(types).toMatch(
			/Extension="svg"[^>]*ContentType="image\/svg\+xml"|ContentType="image\/svg\+xml"[^>]*Extension="svg"/u,
		);
		expect(pictures[0]?.svgPath).toMatch(/\.svg$/u);
	});

	it('writes the extension list last in the blip', async () => {
		const { xml } = await save([[SVG]]);
		const [blip = ''] = blips(xml);
		expect(blip).toMatch(/<a:extLst>[\s\S]*<\/a:extLst><\/a:blip>$/u);
	});

	it('writes the effects of an SVG picture before its extension list', async () => {
		const { xml, pictures } = await save([[SVG, { opacity: 0.5, brightness: 20, contrast: -10 }]]);
		const [blip = ''] = blips(xml);
		const extLst = blip.indexOf('<a:extLst>');
		expect(blip).toMatch(/<a:extLst>[\s\S]*<\/a:extLst><\/a:blip>$/u);
		for (const effect of [
			'<a:alphaModFix amt="50000"',
			'<a:lum bright="20000" contrast="-10000"',
		]) {
			expect(blip.indexOf(effect)).toBeGreaterThan(0);
			expect(blip.indexOf(effect)).toBeLessThan(extLst);
		}
		expect(svgEmbed(blip)).not.toBe('');
		expect(pictures[0]?.svgPath).toMatch(/\.svg$/u);
		expect(pictures[0]?.imageEffects?.alphaModFix).toBe(50);
		expect(pictures[0]?.imageEffects?.lum).toEqual({ bright: 20, contrast: -10 });
	});

	it('uses the given raster fallback', async () => {
		const { zip } = await save([[SVG, { svgFallback: PNG }]]);
		const png = Object.keys(zip.files).find((name) => name.endsWith('.png')) ?? '';
		const bytes = await zip.file(png)?.async('base64');
		expect(PNG.endsWith(bytes ?? '-')).toBe(true);
	});

	it('reads a percent-encoded or plain-text SVG data URL', async () => {
		for (const source of [
			`data:image/svg+xml,${encodeURIComponent(SVG_TEXT)}`,
			`data:image/svg+xml;charset=utf-8,${encodeURIComponent(SVG_TEXT)}`,
		]) {
			const { zip, xml, target } = await save([[source]]);
			const svgTarget = target(svgEmbed(blips(xml)[0] ?? '')) ?? '';
			const svgPart = await zip.file(`ppt/${svgTarget.replace(/^\.\.\//u, '')}`)?.async('string');
			expect(svgPart).toBe(SVG_TEXT);
		}
	});

	it('shares one SVG part and one raster part between identical SVG pictures', async () => {
		const { zip, xml, pictures } = await save([[SVG], [SVG, { x: 300 }]]);
		const media = Object.keys(zip.files).filter((name) => name.startsWith('ppt/media/'));
		expect(media).toHaveLength(2);
		const [a = '', b = ''] = blips(xml);
		expect(svgEmbed(a)).toBe(svgEmbed(b));
		expect(pictures.map((p) => p.svgPath)).toEqual([pictures[0]?.svgPath, pictures[0]?.svgPath]);
	});

	it('leaves raster pictures without the extension', async () => {
		const { xml } = await save([[PNG]]);
		expect(xml).not.toContain('svgBlip');
	});

	it('rejects a fallback it cannot use', () => {
		expect(() => createImageElement(SVG, { svgFallback: SVG })).toThrow(RangeError);
		expect(() => createImageElement(PNG, { svgFallback: PNG })).toThrow(RangeError);
		expect(() => createImageElement('data:image/svg+xml,not-svg')).toThrow(RangeError);
	});
});
