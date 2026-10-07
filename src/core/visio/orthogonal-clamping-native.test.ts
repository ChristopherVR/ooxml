import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';

// Generate locally using scripts/verify-visio-rounding.ps1. Native exports are
// external evidence, not distributed fixtures or a requirement to install Visio in CI.
const directory = process.env.VISIO_NATIVE_ROUNDING_DIR;
describe.skipIf(!directory)('native Visio rounding package', () => {
	it('parses nine native-authored shapes with the same radii as native SVG export', async () => {
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			Version: string;
			Cases: { ShapeId: string; Radii: number[]; NativePath: string }[];
		};
		expect(evidence.Version).toBe('16.0');
		expect(evidence.Cases).toHaveLength(9);
		const document = await parseVsdx(await readFile(join(directory!, 'rounding.vsdx')));
		for (const sample of evidence.Cases) {
			const shape = document.pages[0]!.shapes.find((s) => s.id === sample.ShapeId)!;
			expect(shape).toBeDefined();
			const native = [...sample.NativePath.matchAll(/A([\d.]+)\s+([\d.]+)/g)];
			const rendered = [...shape.geometry[0]!.path.matchAll(/A ([\d.]+) ([\d.]+)/g)];
			expect(rendered).toHaveLength(sample.Radii.length);
			expect(native).toHaveLength(sample.Radii.length);
			for (let i = 0; i < rendered.length; i++) {
				for (const axis of [1, 2]) {
					expect(Number(rendered[i]![axis])).toBeCloseTo(sample.Radii[i]!, 8);
					expect(Number(rendered[i]![axis])).toBeCloseTo(Number(native[i]![axis]) / 72, 6);
				}
			}
			const endpoints = (path: string) =>
				[...path.matchAll(/([MLA])\s*([^MLA]+)/g)].map((command) => {
					const values = command[2]!.trim().split(/\s+/).map(Number);
					return {
						type: command[1],
						x: values.at(-2)!,
						y: values.at(-1)!,
						sweep: command[1] === 'A' ? values[4] : undefined,
					};
				});
			const actualPoints = endpoints(shape.geometry[0]!.path);
			const nativePoints = endpoints(sample.NativePath);
			expect(actualPoints).toHaveLength(nativePoints.length);
			for (let i = 0; i < actualPoints.length; i++) {
				const actual = actualPoints[i]!,
					reference = nativePoints[i]!;
				expect(actual.type).toBe(reference.type);
				// Native SVG uses points and a downward Y axis; align export origins.
				expect(
					Math.abs(actual.x - actualPoints[0]!.x - (reference.x - nativePoints[0]!.x) / 72),
				).toBeLessThan(0.0002);
				expect(
					Math.abs(actual.y - actualPoints[0]!.y + (reference.y - nativePoints[0]!.y) / 72),
				).toBeLessThan(0.0002);
				if (actual.type === 'A') expect(actual.sweep).toBe(1 - reference.sweep!);
			}
			expect(
				document.diagnostics.filter(
					(d) => d.shapeId === sample.ShapeId && d.code === 'unsupported-corner-rounding',
				),
			).toEqual([]);
		}
	});
});
