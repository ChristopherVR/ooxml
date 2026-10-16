import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { loadThemeVariants } from './theme-variants';

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
const parse = (xml: string) => parser.parse(xml) as Record<string, unknown>;

function theme(accent1: string, dk1: string, lt1: string): string {
	const slot = (key: string, val: string) => `<a:${key}><a:srgbClr val="${val}"/></a:${key}>`;
	return (
		'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Facet"><a:themeElements>' +
		`<a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="${dk1}"/></a:dk1>${slot('lt1', lt1)}` +
		`${slot('dk2', '2C3C43')}${slot('lt2', 'EBEBEB')}${slot('accent1', accent1)}${slot('accent2', '54A021')}` +
		`${slot('accent3', 'E6B91E')}${slot('accent4', 'E76618')}${slot('accent5', 'C42F1A')}${slot('accent6', '918655')}` +
		`${slot('hlink', '9AB5E4')}${slot('folHlink', 'C0A0F0')}</a:clrScheme>` +
		'<a:fontScheme name="x"><a:majorFont><a:latin typeface="Trebuchet MS"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
		'<a:minorFont><a:latin typeface="Trebuchet MS"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
		'<a:fmtScheme name="Facet"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>' +
		'</a:themeElements></a:theme>'
	);
}

function master(bg: string, map: string): string {
	return (
		'<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
		`<p:cSld>${bg}<p:spTree/></p:cSld><p:clrMap ${map}/></p:sldMaster>`
	);
}

const LIGHT = 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"';
const DARK = 'bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"';
const REF = (idx: number, scheme: string) =>
	`<p:bg><p:bgRef idx="${idx}"><a:schemeClr val="${scheme}"/></p:bgRef></p:bg>`;

async function variantPackage(): Promise<JSZip> {
	const zip = new JSZip();
	const rels = (items: string) =>
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
	zip.file(
		'_rels/.rels',
		rels(
			'<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/themeVariants" Target="/themeVariants/themeVariantManager.xml"/>',
		),
	);
	zip.file(
		'themeVariants/themeVariantManager.xml',
		'<t:themeVariantManager xmlns:t="http://schemas.microsoft.com/office/thememl/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><t:themeVariantLst>' +
			'<t:themeVariant name="Facet" vid="{A}" cx="12192000" cy="6858000" r:id="rId1"/>' +
			'<t:themeVariant name="Facet" vid="{B}" cx="12192000" cy="6858000" r:id="rId2"/>' +
			'<t:themeVariant name="Facet" vid="{A}" cx="9144000" cy="6858000" r:id="rId3"/>' +
			'</t:themeVariantLst></t:themeVariantManager>',
	);
	zip.file(
		'themeVariants/_rels/themeVariantManager.xml.rels',
		rels(
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/theme/theme/themeManager.xml"/>' +
				'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/themeVariants/variant1/theme/theme/themeManager.xml"/>' +
				'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/themeVariants/variant2/theme/theme/themeManager.xml"/>',
		),
	);
	zip.file('theme/theme/theme1.xml', theme('90C226', '000000', 'FFFFFF'));
	zip.file('theme/slideMasters/slideMaster1.xml', master(REF(1001, 'bg1'), LIGHT));
	zip.file('themeVariants/variant1/theme/theme/theme1.xml', theme('CC0000', '000000', 'FFFFFF'));
	zip.file(
		'themeVariants/variant1/theme/slideMasters/slideMaster1.xml',
		master(REF(1002, 'bg1'), DARK),
	);
	zip.file('themeVariants/variant2/theme/theme/theme1.xml', theme('0000CC', '000000', 'FFFFFF'));
	return zip;
}

describe('loadThemeVariants', () => {
	it('returns nothing for a package without a themeVariants relationship', async () => {
		expect(await loadThemeVariants(new JSZip(), parse as never)).toEqual([]);
	});

	it('reads each variant of the deck slide size, theme and master', async () => {
		const variants = await loadThemeVariants(await variantPackage(), parse as never, 12192000);
		expect(variants.map((variant) => variant.id)).toEqual(['{A}', '{B}']);
		expect(variants[0]!.colorScheme.accent1).toBe('#90C226');
		expect(variants[1]!.colorScheme.accent1).toBe('#CC0000');
		expect(variants[1]!.colorScheme.dk1).toBe('#000000');
		expect(variants[1]!.fontScheme.majorFont?.latin).toBe('Trebuchet MS');
		expect(variants[0]!.backgroundStyleIndex).toBe(1);
		// Style 8 is the dark bg1 column of the second row.
		expect(variants[1]!.backgroundStyleIndex).toBe(8);
		expect(variants[1]!.formatSchemeXml).toMatch(/^<a:fmtScheme name="Facet">/u);
	});

	it('falls back to the first size group when none matches the deck', async () => {
		const variants = await loadThemeVariants(await variantPackage(), parse as never, 1);
		expect(variants.map((variant) => variant.id)).toEqual(['{A}', '{B}']);
	});

	it('uses the 4:3 group for a 4:3 deck', async () => {
		const variants = await loadThemeVariants(await variantPackage(), parse as never, 9144000);
		expect(variants.map((variant) => variant.colorScheme.accent1)).toEqual(['#0000CC']);
	});
});
