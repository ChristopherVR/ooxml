import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { parseXml, buildXml } from '../xml/index';
import { VISIO_NS } from './sheet';
import { recalculateVisioCells, visioCellDependsOn } from './edit-recalculate';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { masterCells } from './edit-master-index';

const root = (extra = '') =>
	parseXml(`<PageContents xmlns="${VISIO_NS}"><Shapes>
<Shape ID="1"><Cell N="Width" V="8"/><Cell N="Height" V="2"/>
<Section N="Property"><Row N="Offset"><Cell N="Value" V="1" F="Width/1IN*0.25"/>
<Cell N="Type" V="2"/></Row></Section>${extra}</Shape>
<Shape ID="2"><Cell N="PinX" V="1" F="Sheet.1!Prop.Offset*1IN"/></Shape>
</Shapes></PageContents>`).documentElement;
const changed = [{ pageId: '0', shapeId: '1', cell: 'Width' }];

describe('native Property section dependency names', () => {
	it('uses the same property names in inherited master scope', () => {
		const page = root();
		const sheet = page.getElementsByTagNameNS(VISIO_NS, 'Shape')[0]!;
		const indexed = masterCells(sheet);
		expect(indexed.get('prop.offset')).toBe(indexed.get('prop.offset.value'));
		expect(indexed.get('prop.offset')?.getAttribute('F')).toBe('Width/1IN*0.25');
	});
	it('updates native Visio property caches through resize and save/reopen', async () => {
		const bytes = new Uint8Array(
			await readFile(new URL('./__fixtures__/native-property-dependencies.vsdx', import.meta.url)),
		);
		const source = (await parseVsdx(bytes)).pages[0]!;
		const edit = await editVsdx(bytes, [
			{ type: 'resize-shape', pageId: source.id, shapeId: '1', width: 8, height: 2 },
		]);
		const saved = (await parseVsdx(edit.bytes)).pages[0]!;
		expect(saved.shapes[0]!.shapeData?.find((row) => row.name === 'Offset')?.value).toBe(2);
	});
	it('recalculates Shape Data and its transitive cross-shape dependents', () => {
		const page = root();
		const roots = new Map([['0', page]]);
		expect(visioCellDependsOn(roots, { pageId: '0', shapeId: '2', cell: 'PinX' }, changed)).toBe(
			true,
		);
		expect(recalculateVisioCells(roots, changed)).toEqual(['0']);
		const cells = Array.from(page.getElementsByTagNameNS(VISIO_NS, 'Cell'));
		expect(cells.find((c) => c.getAttribute('N') === 'Value')?.getAttribute('V')).toBe('2');
		expect(cells.find((c) => c.getAttribute('N') === 'PinX')?.getAttribute('V')).toBe('2');
		expect(cells.find((c) => c.getAttribute('N') === 'Value')?.getAttribute('F')).toBe(
			'Width/1IN*0.25',
		);
	});
	it('rejects ambiguous Prop and Property aliases atomically', () => {
		const page = root('<Section N="Prop"><Row N="Offset"><Cell N="Value" V="3"/></Row></Section>');
		const before = buildXml(page.ownerDocument!);
		expect(() => recalculateVisioCells(new Map([['0', page]]), changed)).toThrow(/Duplicate/);
		expect(buildXml(page.ownerDocument!)).toBe(before);
	});
	it('retains atomic refusal for unsupported affected Shape Data formulas', () => {
		const page = root();
		Array.from(page.getElementsByTagNameNS(VISIO_NS, 'Cell'))
			.find((c) => c.getAttribute('N') === 'Value')!
			.setAttribute('F', 'MOD(Width/1IN,0)');
		const before = buildXml(page.ownerDocument!);
		expect(() => recalculateVisioCells(new Map([['0', page]]), changed)).toThrow(
			/unsupported functions/,
		);
		expect(buildXml(page.ownerDocument!)).toBe(before);
	});
});
