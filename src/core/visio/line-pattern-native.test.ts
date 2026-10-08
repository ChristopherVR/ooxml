import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { visioLineDashLengths } from './ui/line-dash';

const directory = process.env.VISIO_NATIVE_LINE_PATTERNS_DIR;
interface NativeLine {
	pattern: number;
	cap: number;
	weight: number;
	scale?: number;
	pageId: string | number;
	shapeId: string | number;
	svgStyle: string;
}
it.skipIf(!directory)(
	'matches native pattern lengths and caps across stroke weights and drawing scales',
	async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim(),
		) as {
			lines: NativeLine[];
		};
		const document = await parseVsdx(
			new Uint8Array(await readFile(join(directory!, 'paint.vsdx'))),
		);
		expect(evidence.lines.length).toBeGreaterThanOrEqual(144);
		for (const entry of evidence.lines) {
			const page = document.pages.find((candidate) => candidate.id === String(entry.pageId))!;
			const shape = page.shapes.find((candidate) => candidate.id === String(entry.shapeId))!;
			expect(shape.style.linePattern).toBe(entry.pattern);
			expect(shape.style.lineCap).toBe(['round', 'butt', 'square'][entry.cap]);
			expect(shape.style.lineWidth).toBeCloseTo(entry.weight / 72, 12);
			const nativeDash = /stroke-dasharray:([^;}]+)/u
				.exec(entry.svgStyle)?.[1]
				?.split(',')
				.map(Number);
			const lengths = visioLineDashLengths(shape.style)?.map((length) => length * 72);
			if (nativeDash) {
				expect(
					lengths,
					`pattern ${entry.pattern}, cap ${entry.cap}, weight ${entry.weight}, scale ${entry.scale ?? 1}`,
				).toHaveLength(nativeDash.length);
				for (let index = 0; index < nativeDash.length; index++)
					expect(lengths![index]).toBeCloseTo(nativeDash[index]!, 9);
			} else expect(lengths).toBeUndefined();
		}
		expect(
			document.diagnostics.some(
				({ code }) => code === 'inferred-line-pattern' || code === 'inferred-line-cap',
			),
		).toBe(false);
	},
	120_000,
);
