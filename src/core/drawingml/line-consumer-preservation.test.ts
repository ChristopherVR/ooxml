import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS, buildXml, elements, parseXml, type XmlElement } from '../xml/index';
import { parseChartSpace } from '../chart/parse-space';
import { setChartTitle } from '../chart/edit/title';
import { writeChartSpace } from '../chart/write-space';
import { parseDrawingLine } from './drawing-fill';

function outlines(root: XmlElement): XmlElement[] {
	return elements(root).flatMap((child) => [
		...(child.namespaceURI === NS.a && child.localName === 'ln' ? [child] : []),
		...outlines(child),
	]);
}

function canonical(node: XmlElement): unknown {
	return [
		node.namespaceURI,
		node.localName,
		Array.from(node.attributes)
			.filter((attr) => attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
			.map((attr) => [attr.namespaceURI, attr.localName, attr.value])
			.sort(),
		elements(node).map(canonical),
	];
}

describe('shared outline fields survive real chart consumer saves', () => {
	async function nativeChart(): Promise<string> {
		const zip = await JSZip.loadAsync(
			readFileSync(path.join(import.meta.dirname, '../xlsx/__fixtures__/excel-features.xlsx')),
		);
		return zip.file('xl/charts/chart1.xml')!.async('string');
	}
	it('keeps every authored native outline on an independent chart-title edit', async () => {
		const xml = await nativeChart();
		const source = outlines(parseXml(xml).documentElement);
		expect(source.length).toBeGreaterThan(0);
		expect(source.some((node) => parseDrawingLine(node)?.join)).toBe(true);
		const saved = writeChartSpace(setChartTitle(parseChartSpace(xml).chartSpace, 'Edited title'));
		expect(outlines(parseXml(saved).documentElement).map(canonical)).toEqual(source.map(canonical));
	});
	it('retains custom dashes, joins, ends, theme tokens and unknown XML through the existing chart writer', async () => {
		const doc = parseXml(await nativeChart());
		const old = outlines(doc.documentElement)[0]!;
		const line = parseXml(
			`<d:ln xmlns:d="${NS.a}" xmlns:v="urn:vendor" w="012700" cmpd="dbl" algn="in" v:flag="keep"><d:solidFill><d:schemeClr val="accent3"><d:alpha val="40000"/></d:schemeClr></d:solidFill><d:custDash><d:ds d="125000" sp="75000" v:stop="keep"/></d:custDash><d:miter lim="800000"/><d:headEnd type="triangle" w="lg" len="sm"/><d:tailEnd type="oval"/><d:extLst><d:ext uri="keep"><v:data/></d:ext></d:extLst></d:ln>`,
		).documentElement;
		old.parentNode!.replaceChild(doc.importNode(line, true), old);
		const xml = buildXml(doc);
		const saved = writeChartSpace(setChartTitle(parseChartSpace(xml).chartSpace, 'Changed'));
		const reopened = outlines(parseXml(saved).documentElement)[0]!;
		expect(canonical(reopened)).toEqual(canonical(line));
		expect(parseDrawingLine(reopened)).toMatchObject({
			compound: 'dbl',
			alignment: 'in',
			customDash: [{ dash: 125000, space: 75000 }],
			join: { kind: 'miter', limit: 800000 },
			headEnd: { type: 'triangle', width: 'lg', length: 'sm' },
			tailEnd: { type: 'oval' },
		});
	});
});
