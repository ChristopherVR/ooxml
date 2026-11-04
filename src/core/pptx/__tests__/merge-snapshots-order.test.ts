import { readFileSync, readdirSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { validateXML } from 'xmllint-wasm';

import { mergePptxSnapshots } from '../merge-snapshots';
import { NS } from '../../xml/namespaces';
import { buildXml, elements, parseXml } from '../../xml/xml';

const fixture = readFileSync(new URL('./fixtures/slide-size-centering.pptx', import.meta.url));
const part = 'ppt/slides/slide1.xml';
const schemaDir = new URL('../../../../schemas/ecma-376-transitional/', import.meta.url);
const schemas = readdirSync(schemaDir)
	.filter((name) => name.endsWith('.xsd'))
	.map((fileName) => ({
		fileName,
		contents: readFileSync(new URL(fileName, schemaDir), 'utf8')
			.replace(
				/schemaLocation="http:\/\/www\.w3\.org\/(?:2001|2009\/01)\/xml\.xsd"/g,
				'schemaLocation="xml.xsd"',
			)
			.replace(
				/<xsd:import namespace="http:\/\/www\.w3\.org\/XML\/1998\/namespace"\s*\/>/g,
				'<xsd:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>',
			),
	}));

function properties(xml: string) {
	const shape = parseXml(xml).getElementsByTagNameNS(NS.p, 'sp')[0]!;
	return elements(shape).find((child) => child.localName === 'spPr')!;
}

async function deck(children: string, prefix = 'a'): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(fixture);
	const doc = parseXml(await zip.file(part)!.async('string'));
	const shape = doc.getElementsByTagNameNS(NS.p, 'sp')[0]!;
	const spPr = elements(shape).find((child) => child.localName === 'spPr')!;
	for (const child of elements(spPr)) {
		if (!['xfrm', 'prstGeom', 'custGeom'].includes(child.localName)) spPr.removeChild(child);
	}
	const fragment = parseXml(
		'<root xmlns:a="' +
			NS.a +
			'" xmlns:d="' +
			NS.a +
			'">' +
			children.replace(/a:/g, prefix + ':') +
			'</root>',
	);
	for (const child of elements(fragment.documentElement))
		spPr.appendChild(doc.importNode(child, true));
	zip.file(part, buildXml(doc));
	return zip.generateAsync({ type: 'uint8array' });
}

async function shapeXml(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return buildXml(properties(await zip.file(part)!.async('string')));
}

async function valid(xml: string): Promise<boolean> {
	const result = await validateXML({
		xml: [{ fileName: 'shape.xml', contents: xml }],
		schema: {
			fileName: 'shape.xsd',
			contents:
				'<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:a="' +
				NS.a +
				'" targetNamespace="' +
				NS.p +
				'" elementFormDefault="qualified">' +
				'<xsd:import namespace="' +
				NS.a +
				'" schemaLocation="dml-main.xsd"/>' +
				'<xsd:element name="spPr" type="a:CT_ShapeProperties"/></xsd:schema>',
		},
		preload: schemas,
	});
	return result.valid;
}

const fill = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>';
const line = '<a:ln w="38100"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>';

describe('PPTX snapshot schema order', () => {
	it.each(['a', 'd'])(
		'merges fill and outline edits in either peer order (%s prefix)',
		async (prefix) => {
			const base = await deck('');
			const local = await deck(fill, prefix);
			const remote = await deck(line, prefix);
			for (const bytes of [base, local, remote])
				expect(await valid(await shapeXml(bytes))).toBe(true);
			for (const [left, right] of [
				[local, remote],
				[remote, local],
			]) {
				const result = await mergePptxSnapshots(base, left!, right!);
				expect(result.status).toBe('merged');
				if (result.status !== 'merged') throw new Error(result.issues.join('\n'));
				const xml = await shapeXml(result.bytes);
				expect(await valid(xml)).toBe(true);
				const names = elements(parseXml(xml).documentElement).map((child) => child.localName);
				expect(names.indexOf('solidFill')).toBeLessThan(names.indexOf('ln'));
				expect(xml).toContain('FF0000');
				expect(xml).toContain('0000FF');
			}
		},
	);

	it('keeps independently added line fill and dash in schema order', async () => {
		const base = await deck('<a:ln/>');
		const left = await deck('<a:ln>' + fill + '</a:ln>');
		const right = await deck('<a:ln><a:prstDash val="dash"/></a:ln>');
		for (const [local, remote] of [
			[left, right],
			[right, left],
		]) {
			const result = await mergePptxSnapshots(base, local!, remote!);
			expect(result.status).toBe('merged');
			if (result.status !== 'merged') throw new Error(result.issues.join('\n'));
			expect(await valid(await shapeXml(result.bytes))).toBe(true);
		}
	});

	it('declines concurrent alternatives of one fill choice', async () => {
		const base = await deck('');
		const left = await deck('<a:noFill/>');
		const right = await deck(fill);
		for (const [local, remote] of [
			[left, right],
			[right, left],
		]) {
			expect((await mergePptxSnapshots(base, local!, remote!)).status).toBe('unresolved');
		}
	});

	it('merges independent effect entries without inventing a new order', async () => {
		const result = await mergePptxSnapshots(
			await deck('<a:effectLst/>'),
			await deck('<a:effectLst><a:softEdge rad="9525"/></a:effectLst>'),
			await deck('<a:effectLst><a:blur rad="9525"/></a:effectLst>'),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') throw new Error(result.issues.join('\n'));
		const xml = await shapeXml(result.bytes);
		expect(await valid(xml)).toBe(true);
		expect(xml).toContain('softEdge');
		expect(xml).toContain('blur');
	});
});
