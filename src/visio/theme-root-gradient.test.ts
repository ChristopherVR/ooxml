import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import { parseVsdx } from './parser.js';
import { readSheet, type Sheet } from './sheet.js';
import { rootThemeSheet } from './theme-root.js';
import { cell, fixture, rectangle, row, section, shape, xml } from './test-fixtures.js';

const gradient = (contents: string) => section('FillGradient', contents);
const stop = (
	index: number,
	color: string,
	transparency: string | number,
	position: string | number,
) =>
	row(
		index,
		'',
		cell('GradientStopColor', color) +
			cell('GradientStopColorTrans', transparency) +
			cell('GradientStopPosition', position),
	);
const read = (contents: string): Sheet =>
	readSheet(parseXml(xml('Shape', contents)).documentElement);
const savedRows = stop(0, '#ff0000', 0.2, 0) + stop(1, '#0000ff', 0.8, 1);
const themedRows = stop(0, 'Themed', 'Themed', 'Themed') + stop(1, 'Themed', 'Themed', 'Themed');
const rows = (sheet: Sheet) => sheet.sections.get('FillGradient:0')!.rows;

describe('explicit root gradient stop caches', () => {
	it('resolves row colors and formats only for explicit root color selection, without mutation', () => {
		const root = read(gradient(savedRows));
		const sheet = read(cell('ColorSchemeIndex', 0) + gradient(themedRows));
		const result = rootThemeSheet(sheet, { rootSheet: root });
		expect(rows(result).get('1')?.cells.get('GradientStopColor')?.value).toBe('#0000ff');
		expect(rows(result).get('1')?.cells.get('GradientStopPosition')?.value).toBe('1');
		expect(rows(result).get('1')?.cells.get('GradientStopColorTrans')?.value).toBe('0.8');
		expect(rows(sheet).get('1')?.cells.get('GradientStopColor')?.value).toBe('Themed');
		expect(rows(root).get('1')?.cells.get('GradientStopColor')?.value).toBe('#0000ff');
	});
	it('root fill matrix selects row formats independently of colors', () => {
		const sheet = read(cell('QuickStyleFillMatrix', 0) + gradient(themedRows));
		const result = rootThemeSheet(sheet, { rootSheet: read(gradient(savedRows)) });
		expect(rows(result).get('0')?.cells.get('GradientStopColor')?.value).toBe('Themed');
		expect(rows(result).get('0')?.cells.get('GradientStopColorTrans')?.value).toBe('0.2');
	});
	it('matches row indices without replacing literal caches or adding absent cells/rows', () => {
		const root = read(gradient(savedRows));
		const sheet = read(
			cell('ColorSchemeIndex', 0) +
				gradient(
					row(1, '', cell('GradientStopColorTrans', 'Themed') + cell('GradientStopPosition', 0.4)),
				),
		);
		const result = rootThemeSheet(sheet, { rootSheet: root });
		expect(rows(result).size).toBe(1);
		expect(rows(result).get('1')?.cells.has('GradientStopColor')).toBe(false);
		expect(rows(result).get('1')?.cells.get('GradientStopPosition')?.value).toBe('0.4');
		expect(rows(result).get('1')?.cells.get('GradientStopColorTrans')?.value).toBe('0.8');
	});
	it.each([
		gradient(savedRows).replace('N="FillGradient"', 'N="FillGradient" Del="1"'),
		gradient(savedRows).replaceAll('<Row ', '<Row Del="1" '),
		gradient(themedRows),
		'',
	])('does not substitute missing, deleted or unresolved root data: %s', (source) => {
		const sheet = read(cell('ColorSchemeIndex', 0) + gradient(themedRows));
		expect(rootThemeSheet(sheet, { rootSheet: read(source) })).toBe(sheet);
	});
	it.each(['', 'Themed', 'bad', '65534'])('does not infer a root fill selector %s', (selector) => {
		const sheet = read(
			(selector ? cell('QuickStyleFillMatrix', selector) : '') + gradient(themedRows),
		);
		expect(rootThemeSheet(sheet, { rootSheet: read(gradient(savedRows)) })).toBe(sheet);
	});
	it('respects deleted local rows/sections and absent local sections', () => {
		for (const local of [
			'',
			gradient(themedRows).replace('N="FillGradient"', 'N="FillGradient" Del="1"'),
			gradient(themedRows).replaceAll('<Row ', '<Row Del="1" '),
		]) {
			const sheet = read(cell('ColorSchemeIndex', 0) + local);
			expect(rootThemeSheet(sheet, { rootSheet: read(gradient(savedRows)) })).toBe(sheet);
		}
	});
	it('reports a root stop formula error when its saved result is substituted', () => {
		const root = read(
			gradient(
				savedRows.replace('N="GradientStopColorTrans"', 'N="GradientStopColorTrans" E="#REF!"'),
			),
		);
		const sheet = read(cell('ColorSchemeIndex', 0) + gradient(themedRows));
		const reports: string[] = [];
		rootThemeSheet(sheet, { rootSheet: root }, (code) => reports.push(code));
		expect(reports).toEqual(['cached-cell-error']);
	});
	it('renders a root gradient through inherited fill style with local color overrides', async () => {
		const scalar =
			cell('FillPattern', 1) +
			cell('FillGradientEnabled', 1) +
			cell('FillGradientDir', 0) +
			cell('FillGradientAngle', 0) +
			cell('RotateGradientWithShape', 1) +
			cell('UseGroupGradient', 0);
		const themedScalar = scalar.replaceAll(/V="[^"]*"/g, 'V="Themed"');
		const parsed = await parseVsdx(
			await fixture({
				document: `<StyleSheets><StyleSheet ID="0" NameU="No Style">${scalar}${gradient(savedRows)}</StyleSheet><StyleSheet ID="1" FillStyle="0">${themedScalar}${gradient(themedRows)}</StyleSheet></StyleSheets>`,
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', rectangle + cell('Width', 4) + cell('Height', 2) + cell('ColorSchemeIndex', 0) + gradient(row(1, '', cell('GradientStopColor', '#00ff00'))), 'FillStyle="1"')}</Shapes>`,
					},
				],
			}),
		);
		const style = parsed.pages[0]!.shapes[0]!.style;
		expect(style.fillGradient).toMatchObject({
			start: [0, 1],
			end: [4, 1],
			stops: [
				{ offset: 0, color: '#ff0000', opacity: 0.8 },
				{ offset: 1, color: '#00ff00' },
			],
		});
		expect(style.fillGradient?.stops[1]?.opacity).toBeCloseTo(0.2);
		expect(parsed.diagnostics.some((d) => d.code.includes('gradient'))).toBe(false);
	});
});
