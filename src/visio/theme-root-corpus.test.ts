import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser.js';
import { parseXml } from '../xml/index.js';
import { children, child, readSheet } from './sheet.js';
import type { VisioShape } from './model.js';

// Optional read-only external corpus check. No third-party fixture bytes are redistributed.
// Apache POI 732120980140d5ed64b482c470e0b625cdb1ab15, test-data/diagram/{60489,60973}.vsdx.
// Source: https://github.com/apache/poi/tree/732120980140d5ed64b482c470e0b625cdb1ab15/test-data/diagram
const directory = process.env.VISIO_THEME_CORPUS_DIR;
const files = [
	{
		name: '60489.vsdx',
		hash: '15494702ecb5554c1dfaee2c09ac35e2ea656e2eed7ba4ba0cf1039e6125fbe1',
		shape: '16',
		colors: 1,
	},
	{
		name: '60973.vsdx',
		hash: 'c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03',
		shape: '11',
		colors: 0,
	},
];
const flatten = (shapes: VisioShape[]): VisioShape[] =>
	shapes.flatMap((shape) => [shape, ...flatten(shape.children)]);

describe.skipIf(!directory)('hash-pinned Apache POI root theme caches', () => {
	it.each(files)(
		'resolves saved root line formatting in $name',
		async ({ name, hash, shape, colors }) => {
			const bytes = await readFile(join(directory!, name));
			expect(createHash('sha256').update(bytes).digest('hex')).toBe(hash);
			const zip = await JSZip.loadAsync(bytes);
			const xml = parseXml(await zip.file('visio/document.xml')!.async('string')).documentElement;
			const root = children(child(xml, 'StyleSheets'), 'StyleSheet').find(
				(style) => style.getAttribute('ID') === '0',
			)!;
			expect(root.getAttribute('NameU')).toBe('No Style');
			const saved = readSheet(root).cells;
			expect(saved.get('LineWeight')?.value).toBe('0.01041666666666667');
			expect(saved.get('LineCap')?.value).toBe('0');
			expect(saved.get('LinePattern')?.value).toBe('1');
			expect(saved.get('ColorSchemeIndex')?.value).toBe('0');
			const pages = parseXml(
				await zip.file('visio/pages/pages.xml')!.async('string'),
			).documentElement;
			const page = child(child(pages, 'Page'), 'PageSheet')!;
			expect(page.getAttribute('LineStyle')).toBe('0');
			expect(page.getAttribute('TextStyle')).toBe('0');
			expect(readSheet(page).cells.has('ColorSchemeIndex')).toBe(false);
			const doc = await parseVsdx(bytes, { maxDiagnostics: 100_000 });
			const selected = flatten(doc.pages[0]!.shapes).find((item) => item.id === shape)!;
			expect(selected.style).toMatchObject({
				lineWidth: 0.01041666666666667,
				lineCap: 'round',
				linePattern: 1,
				lineColor: '#000000',
				fill: '#ffffff',
			});
			const codes = doc.diagnostics.map((d) => d.code);
			expect(codes).not.toContain('unresolved-line-cap');
			expect(codes).not.toContain('unresolved-line-pattern');
			expect(codes.filter((code) => code === 'unsupported-color')).toHaveLength(colors);
			// Font/effect/formula gaps are still diagnosed, rather than hidden by root fallback.
			expect(codes).toContain('missing-cached-value');
			if (name === '60489.vsdx') {
				const all = flatten(doc.pages[0]!.shapes);
				expect(all.find((item) => item.id === '2')!.style.lineWidth).toBe(0.006944444444444444);
				expect(all.find((item) => item.id === '111')!.style.fill).toBe('#5b9bd5');
				expect(doc.diagnostics.filter((d) => d.code === 'unsupported-color')).toEqual([
					expect.objectContaining({
						shapeId: '111',
						message: 'Cell Color uses an unresolved color; a default was used.',
					}),
				]);
			}
		},
	);
});
