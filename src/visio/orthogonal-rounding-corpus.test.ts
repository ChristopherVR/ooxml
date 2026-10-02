import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseXml } from '../xml/index.js';
import { children, readSheet, type Sheet } from './sheet.js';
import { parseVsdx } from './parser.js';
// Read-only external fixture, Apache POI 732120980140d5ed64b482c470e0b625cdb1ab15.
const directory = process.env.VISIO_THEME_CORPUS_DIR;
describe.skipIf(!directory)('hash-pinned Apache POI 60973 connector rounding', () => {
	it('rounds fourteen proven corners while retaining two unresolved short-segment connectors', async () => {
		const bytes = await readFile(join(directory!, '60973.vsdx'));
		expect(createHash('sha256').update(bytes).digest('hex')).toBe(
			'c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03',
		);
		const zip = await JSZip.loadAsync(bytes);
		const model = await parseVsdx(bytes, { maxDiagnostics: 100_000 });
		const flatten = (shapes: (typeof model.pages)[0]['shapes']): typeof shapes =>
			shapes.flatMap((s) => [s, ...flatten(s.children)]);
		const radius = 0.07874015748031496;
		const masters = new Map<string, Sheet>();
		for (const master of ['master26', 'master31']) {
			const root = parseXml(
				await zip.file(`visio/masters/${master}.xml`)!.async('string'),
			).documentElement;
			const masterShape = children(children(root, 'Shapes')[0], 'Shape').find(
				(s) => s.getAttribute('ID') === '5',
			)!;
			masters.set(master, readSheet(masterShape));
			expect(Number(readSheet(masterShape).cells.get('Rounding')!.value)).toBe(radius);
		}
		for (const [pageId, part, id, cornerCount] of [
			['4', 'page2', '814', 2],
			['4', 'page2', '830', 2],
			['4', 'page2', '835', 4],
			['4', 'page2', '857', 2],
			['7', 'page3', '293', 4],
			['4', 'page2', '825', 0],
			['7', 'page3', '149', 0],
		] as const) {
			const rendered = flatten(model.pages.find((p) => p.id === pageId)!.shapes).find(
				(s) => s.id === id,
			)!;
			const root = parseXml(
				await zip.file(`visio/pages/${part}.xml`)!.async('string'),
			).documentElement;
			const findShape = (element: Element): Element | undefined => {
				for (const shape of children(element, 'Shape')) {
					if (shape.getAttribute('ID') === id) return shape;
					for (const group of children(shape, 'Shapes')) {
						const found = findShape(group);
						if (found) return found;
					}
				}
				return undefined;
			};
			const raw = readSheet(findShape(children(root, 'Shapes')[0]!)!);
			const geometry = [...raw.sections.values()].find((s) => s.name === 'Geometry')!;
			const masterGeometry = [
				...masters.get(id === '293' ? 'master31' : 'master26')!.sections.values(),
			].find((s) => s.name === 'Geometry')!;
			const points = [...geometry.rows.values()].map((r) =>
				['X', 'Y'].map((axis) => {
					const cached = r.cells.get(axis) ?? masterGeometry.rows.get(r.index)!.cells.get(axis);
					expect(cached).toBeDefined();
					return Number(cached!.value);
				}),
			);
			const path = rendered.geometry[0]!.path;
			const clean = (n: number) => Number(n.toFixed(9));
			expect(path.startsWith(`M ${points[0]!.map(clean).join(' ')} `)).toBe(true);
			expect(path.endsWith(`L ${points.at(-1)!.map(clean).join(' ')}`)).toBe(true);
			expect(path.match(/ A /g) ?? []).toHaveLength(cornerCount);
			expect(rendered.geometry[0]).toMatchObject({ fill: false, stroke: true });
			expect(
				model.diagnostics.some(
					(d) =>
						d.pageId === pageId && d.shapeId === id && d.code === 'unsupported-corner-rounding',
				),
			).toBe(cornerCount === 0);
			if (cornerCount === 0)
				expect(path).toBe(
					points.map((p, i) => `${i ? 'L' : 'M'} ${p.map(clean).join(' ')}`).join(' '),
				);
			if (id === '830') {
				// Independently calculated tangent endpoints, not a renderer snapshot.
				const x = points[1]![0]!,
					y0 = points[1]![1]!,
					y1 = points[2]![1]!;
				expect(y1 - y0).toBeGreaterThan(2 * radius);
				expect(path).toContain(
					`L ${clean(x - radius)} ${clean(y0)} A ${clean(radius)} ${clean(radius)} 0 0 1 ${clean(x)} ${clean(y0 + radius)}`,
				);
				expect(path).toContain(
					`L ${clean(x)} ${clean(y1 - radius)} A ${clean(radius)} ${clean(radius)} 0 0 0 ${clean(x + radius)} ${clean(y1)}`,
				);
			}
		}
	});
});
