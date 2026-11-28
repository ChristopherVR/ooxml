// @vitest-environment node
import { readFileSync } from 'node:fs';
import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { exportDocxPdf } from './index';
import { referenceFonts } from './__fixtures__/word-reference/reference';

/** The decoded content streams of a PDF (fonts and images left out). */
function content(bytes: Uint8Array): string {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return streams.join('\n');
}

async function exported(name: string) {
	const docx = readFileSync(
		new URL(`./__fixtures__/horizontal-rule/${name}.docx`, import.meta.url),
	);
	const { model } = await loadDocx(new Uint8Array(docx));
	const { bytes } = await exportDocxPdf(model, {
		fonts: referenceFonts(),
		now: new Date('2026-11-28T00:00:00Z'),
	});
	const ops = content(bytes);
	// The content is drawn in px, under a 0.75 scale: everything below is in points.
	const baselines = [...ops.matchAll(/Tm\b/g)].length
		? [...ops.matchAll(/1 0 0 -1 [\d.]+ ([\d.]+) Tm/g)].map((match) => Number(match[1]) * 0.75)
		: [];
	const fills = [
		...ops.matchAll(/([\d.]+) ([\d.]+) ([\d.]+) rg\n([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re\nf/g),
	].map((match) => ({
		rgb: [match[1], match[2], match[3]].map(Number),
		x: Number(match[4]) * 0.75,
		y: Number(match[5]) * 0.75,
		width: Number(match[6]) * 0.75,
		height: Number(match[7]) * 0.75,
	}));
	return { baselines: [...new Set(baselines.map((y) => y.toFixed(2)))].map(Number), fills };
}

/**
 * Word 16's "Save as PDF" of a paragraph holding Word's Horizontal Line (`v:rect` with `o:hr="t"`,
 * 1.5 pt, `#a0a0a0`) between two lines of Calibri 11 pt (A4, 2.5 cm margins), measured in the
 * file and its control (empty paragraphs instead of the rules): the baselines are 81.24, 108.14
 * and 135.02 pt in both, and each rule is a 453.6 pt wide rectangle from 70.8 pt that starts 9.37
 * pt below the top of its line (93.60 and 120.50 pt) and ends 2.52 pt above its bottom.
 */
describe('exportDocxPdf of a VML horizontal rule (#435)', async () => {
	const bug = await exported('vml-horizontal-rule');
	const control = await exported('vml-horizontal-rule-control');
	const rules = bug.fills.filter((fill) => fill.width > 400 && fill.height > 1);

	it('keeps the height of the rule paragraph, as an empty paragraph has', () => {
		expect(control.baselines).toEqual(bug.baselines);
		expect(bug.baselines[0]).toBeCloseTo(81.24, 1);
		expect(bug.baselines[1]).toBeCloseTo(108.14, 1);
		expect(bug.baselines[2]).toBeCloseTo(135.02, 1);
	});

	it('draws a 1.5 pt grey rule across the text column where Word draws it', () => {
		expect(rules).toHaveLength(2);
		for (const [index, top] of [93.6, 120.5].entries()) {
			const rule = rules[index]!;
			expect(rule.x).toBeCloseTo(70.8, 0);
			expect(rule.width).toBeCloseTo(453.6, 1);
			expect(Math.abs(rule.y - top)).toBeLessThan(0.15);
			expect(rule.height).toBeCloseTo(1.5, 1);
			expect(rule.rgb[0]).toBeCloseTo(0.627, 2);
		}
		expect(control.fills.filter((fill) => fill.width > 400 && fill.height > 1)).toHaveLength(0);
	});

	it('bevels a shaded rule with the light and dark edges Word draws', () => {
		const first = rules[0]!;
		const light = bug.fills.filter((fill) => fill.rgb[0]! > 0.88 && fill.y >= first.y);
		expect(light.length).toBeGreaterThanOrEqual(2);
		expect(light[0]!.height).toBeCloseTo(0.24, 2);
	});

	it('sizes, aligns and colours a rule by its percentage, alignment and height', async () => {
		// Word: a 6 pt red rule over the left half, and a right aligned grey half-width rule.
		const variants = await exported('vml-horizontal-rule-variants');
		const wide = variants.fills.filter((fill) => fill.width > 200 && fill.height > 1);
		const red = wide.find((fill) => fill.rgb[0] === 1 && fill.rgb[1] === 0)!;
		expect(red.x).toBeCloseTo(70.8, 0);
		expect(red.width).toBeCloseTo(226.8, 1);
		expect(red.height).toBeCloseTo(6, 1);
		expect(Math.abs(red.y - 89.15)).toBeLessThan(0.15);
		const right = wide.find((fill) => fill.rgb[0]! > 0.6 && fill.rgb[0]! < 0.65)!;
		expect(right.x).toBeCloseTo(297.6, 0);
		expect(right.width).toBeCloseTo(226.8, 1);
	});
});
