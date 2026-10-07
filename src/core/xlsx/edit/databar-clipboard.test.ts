import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { elements, NS, parseXml, type XmlElement } from '../../xml/index.js';
import { formatRange, parseRange } from '../address.js';
import type { Worksheet } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from '../write/index.js';
import { createCalcEngine } from '../formula/index.js';
import { createConditionalFormatEvaluator } from '../layout/cf-evaluator.js';
import { createEditSession } from './session.js';

interface NativeCase {
	variant: number;
	source: string;
	transpose: boolean;
	mode: 'all' | 'formats';
	before: string;
	after: string;
}
const native = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-databar-clipboard.json', import.meta.url), 'utf8'),
) as { cases: NativeCase[] };
const range = (ref: string) => parseRange(ref)!;
const seed = saveXlsx(createWorkbook());
async function loadSheet(xml: string) {
	const zip = await JSZip.loadAsync(await seed);
	zip.file('xl/worksheets/sheet1.xml', xml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}
interface XmlShape {
	name: string;
	ns: string | null;
	attributes: [string, string][];
	children: XmlShape[];
	text: string;
}
const shape = (node: XmlElement): XmlShape => ({
	name: node.localName,
	ns: node.namespaceURI,
	attributes: Array.from(node.attributes)
		.filter(
			(a) =>
				a.namespaceURI !== 'http://www.w3.org/2000/xmlns/' &&
				!(node.localName === 'cfRule' && a.name === 'id'),
		)
		.map((a): [string, string] => [a.name, a.value])
		.sort((a, b) => a[0].localeCompare(b[0])),
	children: elements(node).map(shape),
	text: elements(node).length ? '' : (node.textContent ?? '').trim(),
});
const normalized = (sheet: Worksheet) =>
	sheet.conditionalFormats
		.flatMap((format) =>
			format.rules.map((rule) => {
				const {
					extensionId: _id,
					extensionXml,
					...base
				} = rule.type === 'dataBar'
					? rule
					: { ...rule, extensionId: undefined, extensionXml: undefined };
				return {
					ranges: format.ranges.map(formatRange),
					...base,
					extension: extensionXml ? shape(parseXml(extensionXml).documentElement) : null,
				};
			}),
		)
		.sort((a, b) => a.priority - b.priority);

describe('Advanced data-bar clipboard settings recorded in native Excel', () => {
	it.each([
		[20, 80],
		[40, 40],
		[0, 0],
	])(
		'keeps typed lengths %i/%i authoritative through editing, paste and save',
		async (minimum, maximum) => {
			const workbook = await loadSheet(native.cases[0]!.before);
			const sheet = workbook.sheets[0]!;
			const original = structuredClone(sheet.conditionalFormats);
			const format = structuredClone(sheet.conditionalFormats[0]!);
			const rule = format.rules[0]!;
			if (rule.type !== 'dataBar') throw new Error('Expected a data bar');
			rule.minLength = minimum;
			rule.maxLength = maximum;
			const session = createEditSession(workbook);
			session.replaceConditionalFormat(0, 0, format);
			const calc = createCalcEngine(workbook);
			const view = createConditionalFormatEvaluator(workbook, 0, (formula, at) =>
				calc.evaluate(formula, at),
			);
			expect(view.at(2, 0)?.dataBar?.fraction).toBe(minimum / 200);
			session.paste(0, range('D4'), session.copy(0, range('A1:A5')), { mode: 'formats' });
			const bytes = await saveXlsx(workbook);
			const zip = await JSZip.loadAsync(bytes);
			const xml = parseXml(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
			for (const namespace of [NS.x, NS.x14]) {
				const bars = Array.from(xml.getElementsByTagNameNS(namespace, 'dataBar'));
				expect(bars).toHaveLength(2);
				for (const bar of bars) {
					expect(bar.getAttribute('minLength')).toBe(String(minimum));
					expect(bar.getAttribute('maxLength')).toBe(String(maximum));
				}
			}
			const reloaded = await loadXlsx(bytes);
			for (const format of reloaded.sheets[0]!.conditionalFormats)
				expect(format.rules[0]).toMatchObject({ minLength: minimum, maxLength: maximum });
			session.undo();
			session.undo();
			expect(sheet.conditionalFormats).toEqual(original);
			session.redo();
			expect(sheet.conditionalFormats[0]!.rules[0]).toMatchObject({
				minLength: minimum,
				maxLength: maximum,
			});
		},
	);
	it('writes the 10/90 legacy fallback for a full-length linked bar without typed percentages', async () => {
		const workbook = await loadSheet(native.cases[0]!.before);
		const rule = workbook.sheets[0]!.conditionalFormats[0]!.rules[0]!;
		if (rule.type !== 'dataBar') throw new Error('Expected a data bar');
		delete rule.minLength;
		delete rule.maxLength;
		const bytes = await saveXlsx(workbook);
		const zip = await JSZip.loadAsync(bytes);
		const xml = parseXml(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
		const base = xml.getElementsByTagNameNS(NS.x, 'dataBar')[0]!;
		expect([base.getAttribute('minLength'), base.getAttribute('maxLength')]).toEqual(['10', '90']);
		const extended = xml.getElementsByTagNameNS(NS.x14, 'dataBar')[0]!;
		expect([extended.getAttribute('minLength'), extended.getAttribute('maxLength')]).toEqual([
			'0',
			'100',
		]);
		expect((await loadXlsx(bytes)).sheets[0]!.conditionalFormats[0]!.rules[0]).toMatchObject({
			minLength: 0,
			maxLength: 100,
		});
	});
	it('writes typed threshold edits into the linked extension as well as the base rule', async () => {
		const workbook = await loadSheet(native.cases[0]!.before);
		const format = structuredClone(workbook.sheets[0]!.conditionalFormats[0]!);
		const rule = format.rules[0]!;
		if (rule.type !== 'dataBar') throw new Error('Expected a data bar');
		rule.min = { type: 'num', value: '-5' };
		rule.max = { type: 'num', value: '25' };
		createEditSession(workbook).replaceConditionalFormat(0, 0, format);
		const reloaded = await loadXlsx(await saveXlsx(workbook));
		const saved = reloaded.sheets[0]!.conditionalFormats[0]!.rules[0]!;
		expect(saved).toMatchObject({ min: rule.min, max: rule.max });
		if (saved.type !== 'dataBar') throw new Error('Expected a data bar');
		const xml = parseXml(saved.extensionXml!);
		expect(
			Array.from(xml.getElementsByTagNameNS(NS.x14, 'cfvo')).map((n) => [
				n.getAttribute('type'),
				n.textContent,
			]),
		).toEqual([
			['num', '-5'],
			['num', '25'],
		]);
	});
	it.each(native.cases)('variant=$variant $source transpose=$transpose $mode', async (recorded) => {
		const workbook = await loadSheet(recorded.before);
		const expected = await loadSheet(recorded.after);
		const sheet = workbook.sheets[0]!;
		const before = normalized(sheet);
		const session = createEditSession(workbook);
		session.paste(0, range('D4'), session.copy(0, range(recorded.source)), recorded);
		expect(normalized(sheet)).toEqual(normalized(expected.sheets[0]!));
		const ids = sheet.conditionalFormats.flatMap((f) =>
			f.rules.flatMap((r) => (r.type === 'dataBar' ? [r.extensionId] : [])),
		);
		expect(new Set(ids).size).toBe(ids.length);
		const saved = await saveXlsx(workbook);
		const zip = await JSZip.loadAsync(saved);
		const xml = parseXml(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
		expect(xml.getElementsByTagNameNS(NS.x14, 'cfRule')).toHaveLength(2);
		const reloaded = await loadXlsx(saved);
		expect(normalized(reloaded.sheets[0]!)).toEqual(normalized(expected.sheets[0]!));
		session.undo();
		expect(normalized(sheet)).toEqual(before);
		session.redo();
		expect(normalized(sheet)).toEqual(normalized(expected.sheets[0]!));
	});
	it('clips, tiles and clears linked extensions without orphan records', async () => {
		const workbook = await loadSheet(native.cases[0]!.before);
		const session = createEditSession(workbook);
		session.paste(0, range('D4:E9'), session.copy(0, range('A2:A4')), 'formats');
		session.clearConditionalFormats(0, range('D4:E4'));
		const saved = await saveXlsx(workbook);
		const reloaded = await loadXlsx(saved);
		expect(reloaded.sheets[0]!.conditionalFormats).toHaveLength(2);
		expect(reloaded.sheets[0]!.conditionalFormats[1]!.ranges.map(formatRange)).toEqual(['D5:E9']);
		session.clearConditionalFormats(0);
		const zip = await JSZip.loadAsync(await saveXlsx(workbook));
		expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).not.toContain('x14:cfRule');
	});
	it('retains unknown bar fields and unrelated worksheet extensions across workbooks', async () => {
		const original = native.cases[0]!.before.replace(
			'<x14:dataBar ',
			'<x14:dataBar future="opaque" ',
		).replace(
			'</extLst></worksheet>',
			'<ext uri="custom"><custom xmlns="urn:test">kept</custom></ext></extLst></worksheet>',
		);
		const workbook = await loadSheet(original);
		const payload = createEditSession(workbook).copy(0, range('A1:A5'));
		const target = createWorkbook();
		createEditSession(target).paste(0, range('D4'), payload);
		const saved = await JSZip.loadAsync(await saveXlsx(target));
		expect(await saved.file('xl/worksheets/sheet1.xml')!.async('string')).toContain(
			'future="opaque"',
		);
		const source = await JSZip.loadAsync(await saveXlsx(workbook));
		expect(await source.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('urn:test');
	});
	it('rewrites formula-valued x14 thresholds with structural edits and cross-sheet cut', async () => {
		const workbook = await loadSheet(native.cases[0]!.before);
		const session = createEditSession(workbook);
		session.insertRows(0, 0, 1);
		const sheet = workbook.sheets[0]!;
		const rule = sheet.conditionalFormats[0]!.rules[0]!;
		expect(rule).toMatchObject({ min: { value: '$J$2' }, max: { value: '$J$3' } });
		if (rule.type !== 'dataBar') throw new Error('Expected a data bar');
		expect(
			parseXml(rule.extensionXml!).getElementsByTagNameNS(
				'http://schemas.microsoft.com/office/excel/2006/main',
				'f',
			)[0]?.textContent,
		).toBe('$J$2');
		const target = session.addSheet('Target');
		session.paste(target, range('D4'), session.cut(0, range('A2:A6')));
		expect(sheet.conditionalFormats).toEqual([]);
		const reloaded = await loadXlsx(await saveXlsx(workbook));
		expect(reloaded.sheets[0]!.conditionalFormats).toEqual([]);
		expect(reloaded.sheets[target]!.conditionalFormats[0]!.ranges.map(formatRange)).toEqual([
			'D4:D8',
		]);
		session.undo();
		expect(sheet.conditionalFormats).toHaveLength(1);
		expect(workbook.sheets[target]!.conditionalFormats).toEqual([]);
	});
});
