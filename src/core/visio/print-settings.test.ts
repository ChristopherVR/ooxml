import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { createVisioPrintSettingsReader, readVisioPrintCountSettings } from './print-settings';
import { visioPrintTileCount } from './print-count';
import { cell, ns } from './test-fixtures';

const refs = 'LineStyle="0" FillStyle="0" TextStyle="0"';
const geometry =
	cell('PageWidth', 17) + cell('PageHeight', 22) + cell('PageScale', 1) + cell('DrawingScale', 1);
const defaults =
	cell('PaperKind', 1) +
	cell('PrintPageOrientation', 1) +
	cell('ScaleX', 1) +
	cell('ScaleY', 1) +
	cell('OnPage', 0) +
	cell('CenterX', 0) +
	cell('CenterY', 0) +
	cell('PagesX', 1) +
	cell('PagesY', 1) +
	['Left', 'Right', 'Top', 'Bottom'].map((side) => cell(`Page${side}Margin`, 0.25)).join('');
const element = (name: string, body: string, attributes = '') =>
	parseXml(`<${name} xmlns="${ns}" ${attributes}>${body}</${name}>`).documentElement;
const read = (local = '', root = defaults, document = '') =>
	readVisioPrintCountSettings(
		[element('PageSheet', geometry + local, refs)],
		[element('DocumentSheet', document, refs)],
		[element('StyleSheet', root, 'ID="0" NameU="No Style"')],
	);
const count = (setup: ReturnType<typeof read>) =>
	visioPrintTileCount({ width: 17, height: 22, pageSetup: setup });

describe('measured No Style0 print defaults', () => {
	it('resolves missing page cells, overrides per cell and ignores DocumentSheet print cells', () => {
		expect(read()).toMatchObject({ paperKind: 1, printFitToPages: false, printZoomY: 1 });
		expect(count(read())).toMatchObject({ kind: 'supported', total: 9 });
		expect(read(cell('PaperKind', 9) + cell('PageLeftMargin', 1))).toMatchObject({
			paperKind: 9,
			margins: { left: 1, right: 0.25 },
		});
		expect(
			read('', defaults, cell('PaperKind', 9) + cell('OnPage', 1) + cell('ScaleX', 0.5)),
		).toEqual(read());
	});
	it.each(['0', 'false', ' 0 ', ' false '])('accepts explicit false cache %j', (value) => {
		expect(count(read(cell('OnPage', value)))).toMatchObject({ kind: 'supported' });
	});
	it.each(['1', 'true', ' 1 ', ' true '])(
		'retains explicit true cache %j as unsupported fit',
		(value) => {
			expect(count(read(cell('OnPage', value)))).toEqual({
				kind: 'unsupported',
				reason: 'fit-mode',
			});
		},
	);
	it.each(['', 'TRUE', '2', '0.0', 'yes'])(
		'refuses malformed local boolean %j without inheriting false',
		(value) => {
			const result = read(cell('OnPage', value));
			expect(result.printFitToPages).toBeUndefined();
			expect(result.printCountRefusal).toBe('invalid-local-setting');
		},
	);
	it.each([
		cell('PaperKind', 'bad'),
		cell('PaperKind', -1),
		cell('PaperKind', 1.5),
		'<Cell N="PaperKind" V="9" E=""/>',
		cell('PaperKind', 9, 'Inh'),
		cell('PaperKind', 9, '8'),
		'<Cell N="PaperKind" V="9" U="IN"/>',
	])('blocks invalid local inheritance: %s', (local) => {
		expect(read(local).printCountRefusal).toBe('invalid-local-setting');
		expect(count(read(local))).toEqual({ kind: 'unsupported', reason: 'invalid-local-setting' });
	});
	it.each([
		cell('ScaleX', 1, 'NOW()'),
		cell('ScaleX', 1, 'PageWidth/17'),
		cell('ScaleX', 1, 'GUARD(1)'),
	])('refuses dynamic or guarded local %s', (local) => {
		expect(read(local).printCountRefusal).toBe('dynamic-setting');
	});
	it('checks duplicate/case-ambiguous cells before maps overwrite them', () => {
		expect(read(cell('PaperKind', 1) + cell('PaperKind', 9)).printCountRefusal).toBe(
			'duplicate-setting',
		);
		expect(read(cell('paperkind', 9)).printCountRefusal).toBe('duplicate-setting');
		expect(read('', defaults + cell('PaperKind', 9)).printCountRefusal).toBe('duplicate-setting');
		expect(read('', defaults, cell('OnPage', 0) + cell('OnPage', 1)).printCountRefusal).toBe(
			'duplicate-setting',
		);
	});
	it('refuses ambiguous sheets and unmeasured root/style chains', () => {
		const page = element('PageSheet', geometry, refs),
			doc = element('DocumentSheet', '', refs);
		const root = element('StyleSheet', defaults, 'ID="0" NameU="No Style"');
		expect(readVisioPrintCountSettings([page, page], [doc], [root]).printCountRefusal).toBe(
			'ambiguous-page-sheet',
		);
		expect(readVisioPrintCountSettings([page], [doc, doc], [root]).printCountRefusal).toBe(
			'ambiguous-document-sheet',
		);
		expect(readVisioPrintCountSettings([page], [doc], [root, root]).printCountRefusal).toBe(
			'duplicate-setting',
		);
		for (const attrs of ['ID="0" NameU="Custom"', 'ID="0" NameU="No Style" LineStyle="0"'])
			expect(
				readVisioPrintCountSettings([page], [doc], [element('StyleSheet', defaults, attrs)])
					.printCountRefusal,
			).toBe('unsupported-print-style');
		page.setAttribute('LineStyle', '3');
		expect(readVisioPrintCountSettings([page], [doc], [root]).printCountRefusal).toBe(
			'unsupported-print-style',
		);
	});
	it('does not invent missing defaults and proves drawing scale without changing its resolution', () => {
		expect(count(read('', defaults.replace(cell('ScaleY', 1), '')))).toEqual({
			kind: 'unsupported',
			reason: 'missing-setting',
		});
		expect(
			readVisioPrintCountSettings(
				[element('PageSheet', cell('PageWidth', 17) + cell('PageHeight', 22), refs)],
				[element('DocumentSheet', '', refs)],
				[element('StyleSheet', defaults, 'ID="0" NameU="No Style"')],
			).printCountRefusal,
		).toBe('unusable-drawing-scale');
	});
	it('indexes a large shared default sheet once for many pages', () => {
		const root = element(
			'StyleSheet',
			defaults + Array.from({ length: 5000 }, (_, i) => cell(`Unrelated${i}`, i)).join(''),
			'ID="0" NameU="No Style"',
		);
		const nodes = root.childNodes;
		let reads = 0;
		Object.defineProperty(root, 'childNodes', {
			get() {
				reads++;
				return nodes;
			},
		});
		const reader = createVisioPrintSettingsReader([element('DocumentSheet', '', refs)], [root]);
		const indexedReads = reads;
		const page = element('PageSheet', geometry, refs);
		for (let i = 0; i < 1000; i++)
			expect(count(reader([page]))).toMatchObject({ kind: 'supported', total: 9 });
		expect(indexedReads).toBeGreaterThan(0);
		expect(reads).toBe(indexedReads);
	});
});
