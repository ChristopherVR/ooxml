// @vitest-environment node
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { describe, expect, it } from 'vitest';
import {
	CHART_PARTS,
	cellFloatCases,
	cellFloatPackage,
} from '../../../../core/docx/__fixtures__/cell-floats/cases';
import { exportDocxPdf } from './export-pdf';

const FIXTURES = join(import.meta.dirname, '../../../../core/docx/__fixtures__');
const measurer = createFakeMeasurer();

/** The decoded page content of a PDF and its standard-font text. */
function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	const content = streams.join('\n');
	const text = [...content.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g)]
		.map((m) => m[1]!.replace(/\\(.)/g, '$1'))
		.join('|');
	return { raw, content, text };
}

describe('exportDocxPdf without a DOM', () => {
	it('runs with no DOM globals', () => {
		expect(globalThis.DOMParser).toBeUndefined();
		expect((globalThis as { document?: unknown }).document).toBeUndefined();
	});

	it('writes a chart as vector paths with real text', async () => {
		const { model } = await loadDocx(new Uint8Array(readFileSync(join(FIXTURES, 'chart.docx'))));
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		expect(pdf.raw).not.toContain('/Subtype /Image');
		expect(pdf.content).toMatch(/ m\n[\d. ]+ l\n[SB]\n/);
		expect(pdf.content).toMatch(/ l\nh\nf\n/);
		expect(pdf.text).toContain('Sales by region');
		expect(pdf.text).toContain('North');
		expect(report.omitted).toEqual([]);
	});

	it('writes a drawing shape and its text box', async () => {
		const chart = await JSZip.loadAsync(readFileSync(join(FIXTURES, 'chart.docx')));
		const parts: Record<string, string> = {};
		for (const part of CHART_PARTS) parts[part] = await chart.file(part)!.async('string');
		const zip = new JSZip();
		for (const [path, xml] of Object.entries(
			cellFloatPackage(cellFloatCases()['textbox-cell']!, parts),
		))
			zip.file(path, xml);
		const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const { bytes, report } = await exportDocxPdf(model, { measurer });
		const pdf = inspect(bytes);
		// The red shape is a filled path; its text is real text.
		expect(pdf.content).toMatch(/1 0 0 rg\n[^]*?\nf/);
		expect(pdf.text).toContain('alpha');
		expect(pdf.text).toContain('beta');
		expect(report.omitted).toEqual([]);
	});
});
