// writeChartSpace against every c:chartSpace part the parser tests read (chart/parse-space-fixtures
// .test.ts): parse -> write -> parse yields the same model, and a part Office or a generator wrote
// with prefixes comes back byte for byte, except for the elements the parser reports as not
// modelled. Those are removed from the source before the comparison, so each difference is pinned.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildXml, elements, isElement, parseXml, type XmlElement } from '../xml/index';
import { parseChartSpace } from './parse-space';
import { writeChartSpace } from './write-space';
import { stripDeclarations } from './xml-fragment';

const root = import.meta.dirname;
const pptx = (name: string) => path.join(root, '../pptx/__tests__/fixtures/e2e', name);
const xlsx = (name: string) => path.join(root, '../xlsx/__fixtures__', name);

/**
 * The elements each part loses, by part: what the parser reports as not modelled. Parts not listed
 * round-trip byte-identical. Causes: trendlines and surface band formats are not in the model yet
 * (docs/agnostic-core-plan.md, step 5).
 */
const UNMODELLED: Record<string, string[]> = {
	'chart-gallery.pptx ppt/charts/chart2.xml': ['c:ser/c:trendline'],
	'three-d-charts.pptx ppt/charts/chart16.xml': ['c:surface3DChart/c:bandFmts'],
	'three-d-charts.pptx ppt/charts/chart17.xml': ['c:surface3DChart/c:bandFmts'],
};

/** Written without prefixes (default chart namespace, `a` declared per element): see below. */
const UNPREFIXED = new Set(['openpyxl-features.xlsx']);

const FILES = [
	pptx('chart-gallery.pptx'),
	pptx('chart-data-fidelity.pptx'),
	pptx('chart-top-axis.pptx'),
	pptx('chart-filtered-series.pptx'),
	pptx('chart-title-runs.pptx'),
	pptx('chart-stacked-line-markers.pptx'),
	pptx('three-d-parity/three-d-charts.pptx'),
	pptx('pie3d.pptx'),
	xlsx('excel-features.xlsx'),
	xlsx('openpyxl-features.xlsx'),
];

async function chartParts(file: string): Promise<[string, string][]> {
	const zip = await JSZip.loadAsync(readFileSync(file));
	const names = Object.keys(zip.files).filter((name) => /charts\/chart\d+\.xml$/.test(name));
	const parts = await Promise.all(
		names.map(async (name) => [name, await zip.file(name)!.async('string')] as [string, string]),
	);
	return parts.filter(([, xml]) => !/<cx:chartSpace/.test(xml));
}

const declarationBreak = (xml: string) =>
	(/^<\?xml[^>]*\?>(\r\n|\n)?/.exec(xml)?.[1] ?? '') as '' | '\n' | '\r\n';

/** The source with the listed `c:parent/c:child` elements removed, root serialised again. */
function withoutElements(xml: string, paths: string[]): string {
	const doc = parseXml(xml);
	const visit = (element: XmlElement) => {
		for (const child of elements(element)) {
			if (paths.includes(`c:${element.localName}/c:${child.localName}`)) element.removeChild(child);
			else visit(child);
		}
	};
	visit(doc.documentElement);
	const declaration = /^<\?xml[^>]*\?>/.exec(xml)?.[0] ?? '';
	return declaration + declarationBreak(xml) + buildXml(doc.documentElement);
}

/** Namespace-aware structure: `{uri}local`, sorted non-declaration attributes, text, children. */
function structure(node: Node): unknown {
	if (!isElement(node)) return node.nodeValue;
	const attributes = Array.from(node.attributes)
		.filter((item) => item.prefix !== 'xmlns' && item.name !== 'xmlns')
		.map((item) => `${item.namespaceURI ?? ''}|${item.localName}=${item.value}`)
		.sort();
	return [
		`{${node.namespaceURI}}${node.localName}`,
		attributes,
		Array.from(node.childNodes).map(structure),
	];
}

describe('writeChartSpace round-trips the parser fixture parts', () => {
	for (const file of FILES) {
		const base = path.basename(file);
		it(`${base}: parse -> write -> parse keeps the model, and the bytes where modelled`, async () => {
			const parts = await chartParts(file);
			expect(parts.length).toBeGreaterThan(0);
			for (const [name, xml] of parts) {
				const first = parseChartSpace(xml);
				const written = writeChartSpace(first.chartSpace, {
					declarationBreak: declarationBreak(xml),
				});
				const second = parseChartSpace(written);
				expect(second.chartSpace, name).toEqual(first.chartSpace);
				expect(second.issues, name).toEqual([]);
				const removed = UNMODELLED[`${base} ${name}`] ?? [];
				expect(new Set(first.issues.map((issue) => issue.message.split(' ')[0])), name).toEqual(
					new Set(removed),
				);
				if (UNPREFIXED.has(base)) {
					// openpyxl writes the chart namespace as the default namespace and declares `a` on
					// each element; the writer uses the c/a/r prefixes Office writes. Same tree otherwise.
					expect(structure(parseXml(written).documentElement), name).toEqual(
						structure(parseXml(xml).documentElement),
					);
				} else if (removed.length) expect(written, name).toBe(withoutElements(xml, removed));
				else expect(written, name).toBe(xml);
			}
		});
	}
});

/** Every chart part embedded in the Excel acceptance fixtures (`"xl/charts/chart1.xml": "<?xml..."`). */
function embeddedParts(value: unknown, out: string[] = []): string[] {
	if (typeof value === 'string') {
		if (value.includes('<c:chartSpace')) out.push(value);
	} else if (value && typeof value === 'object')
		for (const item of Object.values(value)) embeddedParts(item, out);
	return out;
}

/**
 * The two differences the writer is allowed on these parts: a literal CR LF inside `a:t` (XML
 * end-of-line handling turns it into LF when the part is parsed, so it is not recoverable from the
 * DOM) and namespace declarations repeated below the root (the repository's own DOM patcher wrote
 * them on `c:txPr`; the writer declares each prefix once, on the root, as Office does).
 */
function normalised(xml: string): string {
	const rootEnd = xml.indexOf('>', xml.indexOf('<c:chartSpace')) + 1;
	const bindings = new Map(
		[...xml.slice(0, rootEnd).matchAll(/ xmlns:([\w.-]+)="([^"]*)"/g)].map(
			([, prefix, uri]) => [prefix!, uri!] as const,
		),
	);
	return (
		xml.slice(0, rootEnd) + stripDeclarations(xml.slice(rootEnd).replace(/\r\n/g, '\n'), bindings)
	);
}

/** Parts per fixture and how many come back byte-identical before normalisation. */
const EMBEDDED: [string, number, number][] = [
	['../xlsx/__fixtures__/excel-chart-gradient-edits.json', 21, 21],
	['../xlsx/__fixtures__/excel-chart-gradients.json', 3, 3],
	['../xlsx/__fixtures__/excel-chart-palette-edits.json', 32, 32],
	['../xlsx/__fixtures__/excel-chart-spacing.json', 13, 13],
	['../xlsx/__fixtures__/excel-chart-transparency.json', 3, 3],
	['../xlsx/__fixtures__/excel-chart-types.json', 6, 6],
	['../xlsx/__fixtures__/excel-gradient-presets.json', 24, 24],
	['excel-built-in-text-styles.json', 96, 96],
	['excel-chart-legend-layout.json', 8, 8],
	['excel-chart-styles.json', 16, 16],
	['excel-chart-title-layout.json', 7, 7],
	['excel-chart-title-wrap.json', 10, 10],
	// Literal CR LF inside title runs.
	['excel-chart-title-spacing.json', 12, 3],
	['excel-chart-title-text.json', 3, 2],
	// Declarations repeated on c:txPr.
	['excel-chart-text-inheritance.json', 24, 0],
];

describe('writeChartSpace round-trips the Excel acceptance fixture parts', () => {
	for (const [file, count, identical] of EMBEDDED) {
		it(`${path.basename(file)}: ${identical} of ${count} parts byte-identical`, () => {
			const parts = embeddedParts(JSON.parse(readFileSync(path.join(root, file), 'utf8')));
			expect(parts).toHaveLength(count);
			let same = 0;
			for (const xml of parts) {
				const first = parseChartSpace(xml);
				expect(first.issues).toEqual([]);
				const written = writeChartSpace(first.chartSpace, {
					declarationBreak: declarationBreak(xml),
				});
				expect(parseChartSpace(written).chartSpace).toEqual(first.chartSpace);
				expect(written).toBe(normalised(xml));
				if (written === xml) same++;
			}
			expect(same).toBe(identical);
		});
	}
});
