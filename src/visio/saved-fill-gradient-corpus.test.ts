import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseXml } from '../xml/index.js';
import { children, child, readSheet } from './sheet.js';
import { parseVsdx } from './parser.js';
// Apache POI 732120980140d5ed64b482c470e0b625cdb1ab15/test-data/diagram/60973.vsdx.
// Optional read-only check; no third-party fixture bytes are redistributed.
const directory = process.env.VISIO_THEME_CORPUS_DIR;
describe.skipIf(!directory)('hash-pinned saved fill gradient limits', () => {
	it('keeps unverified non-horizontal and radial 60973 gradients diagnosed', async () => {
		const bytes = await readFile(join(directory!, '60973.vsdx'));
		expect(createHash('sha256').update(bytes).digest('hex')).toBe(
			'c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03',
		);
		const model = await parseVsdx(bytes, { maxDiagnostics: 100_000 });
		const warnings = model.diagnostics.filter((d) => d.code === 'unsupported-saved-fill-gradient');
		expect(warnings.map((d) => [d.pageId, d.shapeId])).toEqual(
			['2', '3', '4', '5', '6'].map((id) => ['10', id]),
		);
		const flatten = (shapes: (typeof model.pages)[0]['shapes']): typeof shapes =>
			shapes.flatMap((shape) => [shape, ...flatten(shape.children)]);
		const selected = flatten(model.pages.find((page) => page.id === '10')!.shapes).filter((shape) =>
			['2', '3', '4', '5', '6'].includes(shape.id),
		);
		expect(selected).toHaveLength(5);
		expect(selected.every((shape) => !shape.style.fillGradient)).toBe(true);
		// The same package has a non-horizontal saved linear style. Its existence
		// is not evidence for the angle sign or a native visual-match claim.
		const zip = await JSZip.loadAsync(bytes);
		const document = parseXml(
			await zip.file('visio/document.xml')!.async('string'),
		).documentElement;
		const source = children(child(document, 'StyleSheets'), 'StyleSheet').find(
			(style) => style.getAttribute('ID') === '8',
		)!;
		const sheet = readSheet(source);
		expect(sheet.cells.get('FillGradientEnabled')?.value).toBe('1');
		expect(sheet.cells.get('FillGradientDir')?.value).toBe('0');
		expect(Number(sheet.cells.get('FillGradientAngle')?.value)).toBeCloseTo((3 * Math.PI) / 2);
	});
});
