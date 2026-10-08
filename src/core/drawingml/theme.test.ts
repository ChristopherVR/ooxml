import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS, first, parseXml } from '../xml/index';
import { resolveDrawingColor } from './drawing-color';
import { parseTheme, parseThemeColorMap, parseThemeXml } from './theme';
import {
	resolveSchemeColor,
	themeColorSlotFor,
	themeDrawingColorTheme,
	themeSlotHex,
} from './theme-color';

const core = path.join(import.meta.dirname, '..');

async function part(file: string, name: string): Promise<string> {
	const zip = await JSZip.loadAsync(readFileSync(path.join(core, file)));
	const entry = zip.file(name);
	if (!entry) throw new Error(`${file} has no ${name}`);
	return entry.async('string');
}

describe('parseTheme on real theme parts', () => {
	it('reads the Word theme of the SmartArt fixture', async () => {
		const theme = parseThemeXml(
			await part('docx/__fixtures__/smartart.docx', 'word/theme/theme1.xml'),
		);
		expect(theme.name).toBeTruthy();
		const dk1 = theme.colorScheme.colors.dk1;
		expect(dk1?.kind).toBe('system');
		expect(dk1?.fallback).toMatch(/^[0-9A-F]{6}$/);
		expect(Object.keys(theme.colorScheme.colors)).toHaveLength(12);
		expect(theme.fontScheme.major.latin).toBeTruthy();
		expect(theme.fontScheme.minor.latin).toBeTruthy();
	});

	it('reads the Excel theme with system colours and script fonts', async () => {
		const theme = parseThemeXml(
			await part('xlsx/__fixtures__/excel-features.xlsx', 'xl/theme/theme1.xml'),
		);
		expect(theme.name).toBe('Office Theme');
		expect(theme.colorScheme.name).toBe('Office');
		expect(theme.colorScheme.colors.dk1).toEqual({
			kind: 'system',
			value: 'windowText',
			transforms: [],
			fallback: '000000',
		});
		expect(theme.colorScheme.colors.lt1).toMatchObject({ value: 'window', fallback: 'FFFFFF' });
		expect(theme.colorScheme.colors.accent1).toEqual({
			kind: 'srgb',
			value: '156082',
			transforms: [],
		});
		expect(theme.fontScheme.major.latin).toBe('Aptos Display');
		expect(theme.fontScheme.minor.latin).toBe('Aptos Narrow');
		// Empty `a:ea`/`a:cs` typefaces are omitted, script overrides kept.
		expect(theme.fontScheme.major.eastAsia).toBeUndefined();
		expect(theme.fontScheme.major.scripts.Jpan).toBe('游ゴシック Light');
		expect(theme.fontScheme.minor.scripts.Arab).toBe('Arial');
	});

	it('reads a PowerPoint theme and resolves through the master colour map', async () => {
		const file = 'pptx/__tests__/fixtures/themed-layout-placeholders.pptx';
		const theme = parseThemeXml(await part(file, 'ppt/theme/theme1.xml'));
		expect(theme.name).toBe('Balloons');
		expect(theme.colorScheme.name).toBe('Balloons 8');
		expect(theme.fontScheme.major).toEqual({ latin: 'Verdana', scripts: {} });
		const master = parseXml(await part(file, 'ppt/slideMasters/slideMaster1.xml'));
		const map = parseThemeColorMap(first(master.documentElement, 'clrMap', NS.p));
		expect(map).toMatchObject({ bg1: 'lt1', tx1: 'dk1', accent1: 'accent1' });
		expect(themeSlotHex(resolveSchemeColor(theme.colorScheme, 'tx1', map))).toBe('006699');
		expect(themeSlotHex(resolveSchemeColor(theme.colorScheme, 'bg2', map))).toBe('FFFFCC');
		expect(themeDrawingColorTheme(theme.colorScheme, map).scheme('accent4')).toBe('#005682');
	});
});

describe('theme colour resolution', () => {
	const a = NS.a;
	const theme = parseThemeXml(`<a:themeOverride xmlns:a="${a}">
<a:clrScheme name="S"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:srgbClr val="eeeeee"/></a:lt1>
<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="DDDDDD"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:prstClr val="red"/></a:accent2></a:clrScheme>
<a:fontScheme name="F"><a:majorFont><a:latin typeface="Major"/></a:majorFont><a:minorFont><a:latin typeface=""/></a:minorFont></a:fontScheme>
</a:themeOverride>`);

	it('reads a theme override, whose schemes sit directly under the root', () => {
		expect(theme.colorScheme.name).toBe('S');
		expect(theme.fontScheme).toEqual({
			name: 'F',
			major: { latin: 'Major', scripts: {} },
			minor: { scripts: {} },
		});
	});

	it('maps logical names through the colour map, with the conventional default', () => {
		expect(themeColorSlotFor('bg1')).toBe('lt1');
		expect(themeColorSlotFor('tx2')).toBe('dk2');
		expect(themeColorSlotFor('bg1', { bg1: 'dk1' })).toBe('dk1');
		expect(themeColorSlotFor('accent1', { accent1: 'accent2' })).toBe('accent2');
		expect(themeColorSlotFor('dk1', { bg1: 'dk1' })).toBe('dk1');
		expect(themeColorSlotFor('phClr')).toBeUndefined();
		expect(themeColorSlotFor('constructor')).toBeUndefined();
	});

	it('gives plain hex for srgb and system colours only', () => {
		expect(themeSlotHex(resolveSchemeColor(theme.colorScheme, 'tx1'))).toBe('111111');
		expect(themeSlotHex(resolveSchemeColor(theme.colorScheme, 'bg1'))).toBe('EEEEEE');
		expect(themeSlotHex(resolveSchemeColor(theme.colorScheme, 'accent2'))).toBeUndefined();
		expect(resolveSchemeColor(theme.colorScheme, 'accent2')?.kind).toBe('preset');
		expect(resolveSchemeColor(theme.colorScheme, 'hlink')).toBeUndefined();
	});

	it('feeds the DrawingML colour resolver', () => {
		const colorTheme = themeDrawingColorTheme(theme.colorScheme, { tx1: 'lt1' });
		expect(colorTheme.scheme('tx1')).toBe('#EEEEEE');
		expect(themeDrawingColorTheme(undefined).scheme('accent1')).toBeUndefined();
		const resolved = resolveDrawingColor(
			{ kind: 'scheme', value: 'accent1', transforms: [{ name: 'alpha', value: '50000' }] },
			colorTheme,
		);
		expect(resolved).toMatchObject({ hex: '#4472C4', alpha: 0.5 });
	});

	it('accepts an element as well as a document', () => {
		const doc = parseXml(`<a:theme xmlns:a="${a}" name="T"/>`);
		expect(parseTheme(doc.documentElement)).toEqual(parseTheme(doc));
		expect(parseTheme(doc)).toEqual({
			name: 'T',
			colorScheme: { colors: {} },
			fontScheme: { major: { scripts: {} }, minor: { scripts: {} } },
		});
	});
});
