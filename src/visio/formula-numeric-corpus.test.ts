import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import { children, attribute } from './sheet.js';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula.js';

const directory = process.env['VISIO_EDIT_CORPUS_DIR'];
// Immutable public Apache POI fixtures; hashes match the pinned downloaded corpus.
const fixtures = [
	{
		file: '60973.vsdx',
		sha256: 'c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03',
		part: 'visio/pages/page1.xml',
		shape: '3',
		match: /^SQRT\(/,
		count: 1,
	},
	{
		file: 'test.vsdx',
		sha256: '558c09c518848c98a815c9684f511d5260585fa0c893f2d570b6284d6a9b29a8',
		part: 'visio/masters/master1.xml',
		shape: '5',
		match: /MODULUS/,
		count: 1,
	},
	{
		file: 'test.vsdx',
		sha256: '558c09c518848c98a815c9684f511d5260585fa0c893f2d570b6284d6a9b29a8',
		part: 'visio/masters/master1.xml',
		shape: '6',
		match: /BITAND/,
		count: 4,
	},
	{
		file: 'test.vsdx',
		sha256: '558c09c518848c98a815c9684f511d5260585fa0c893f2d570b6284d6a9b29a8',
		part: 'visio/masters/master1.xml',
		shape: '8',
		match: /^IF\(User.HAlign=1,0,BITOR/,
		count: 1,
	},
	{
		file: 'test_text_extraction.vsdx',
		sha256: '5a5117a6c41313a6e7253689d7ae09c7e46bfbb958a670f60f7f5d0bbaef27f9',
		part: 'visio/masters/master3.xml',
		shape: '5',
		match: /BITXOR|MODULUS/,
		count: 2,
	},
];
function indexedCells(shape: Element) {
	const result = new Map<string, Element>();
	for (const cell of children(shape, 'Cell')) result.set(attribute(cell, 'N')!, cell);
	for (const section of children(shape, 'Section')) {
		const name = attribute(section, 'N')!;
		for (const row of children(section, 'Row'))
			for (const cell of children(row, 'Cell')) {
				const rowName = attribute(row, 'N') ?? attribute(row, 'IX') ?? '0';
				const cellName = attribute(cell, 'N')!;
				if (name === 'User' && cellName === 'Value') result.set(`User.${rowName}`, cell);
			}
	}
	return result;
}
describe.skipIf(!directory)('immutable public numeric ShapeSheet corpus', () => {
	for (const fixture of fixtures)
		it(`evaluates real formulas and caches in ${fixture.file} ${fixture.part} shape ${fixture.shape}`, async () => {
			const source = new Uint8Array(
				readFileSync(join(directory!, 'poi', 'test-data', 'diagram', fixture.file)),
			);
			const snapshot = source.slice();
			expect(createHash('sha256').update(source).digest('hex')).toBe(fixture.sha256);
			const zip = await JSZip.loadAsync(source);
			const originalXml = await zip.file(fixture.part)!.async('string');
			const doc = parseXml(originalXml);
			const shape = Array.from(doc.documentElement.getElementsByTagName('*')).find(
				(node) => node.localName === 'Shape' && attribute(node, 'ID') === fixture.shape,
			)!;
			const cells = indexedCells(shape);
			const owns = (node: Element): boolean => {
				let parent = node.parentNode;
				while (parent && parent !== shape) {
					if (parent.nodeType === 1 && (parent as Element).localName === 'Shape') return false;
					parent = parent.parentNode;
				}
				return parent === shape;
			};
			const targets = Array.from(shape.getElementsByTagName('*')).filter(
				(node) =>
					owns(node) && node.localName === 'Cell' && fixture.match.test(attribute(node, 'F') ?? ''),
			);
			let count = 0;
			for (const target of targets) {
				const formula = attribute(target, 'F')!;
				const analysis = analyzeVisioFormula(formula);
				// These specific source formulas must now be inside the numeric evaluation subset.
				expect(analysis.dynamic, formula).toBe(false);
				expect(analysis.unsupportedFunctions, formula).toEqual([]);
				const value = evaluateVisioFormula(formula, (ref) => {
					expect(ref.shapeId, formula).toBeUndefined();
					const cell = cells.get(ref.cell)!;
					expect(cell, `${formula}: missing ${ref.cell}`).toBeDefined();
					const inferred = /^(Width|Height|TxtWidth|TxtHeight|BeginX|BeginY|EndX|EndY)$/.test(
						ref.cell,
					)
						? 'DL'
						: ref.cell === 'Angle'
							? 'DA'
							: undefined;
					return visioFormulaCachedValue(attribute(cell, 'V')!, attribute(cell, 'U') ?? inferred);
				});
				expect(value.value, formula).toBeCloseTo(Number(attribute(target, 'V')), 8);
				count++;
			}
			expect(count).toBe(fixture.count);
			console.info(`Numeric corpus coverage ${fixture.file}: ${count} source formulas`);
			expect(await zip.file(fixture.part)!.async('string')).toBe(originalXml);
			expect(source).toEqual(snapshot);
		});
});
