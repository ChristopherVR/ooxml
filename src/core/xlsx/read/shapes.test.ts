import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { FormControlObject, UnsupportedDrawingObject } from '../model';
import { createEditSession } from '../edit/index';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './index';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-shapes.xlsx')),
	);

describe('worksheet shapes and form controls', () => {
	it('reads shapes, connectors and groups', async () => {
		const wb = await loadXlsx(fixture());
		const shapes = wb.sheets[0]!.drawings.filter(
			(d): d is UnsupportedDrawingObject => d.kind === 'unsupported' && !!d.shape,
		);
		expect(shapes.map((s) => s.shape!.name)).toEqual([
			'Box',
			'Oval',
			'Arrow',
			'Note',
			'Line',
			'Pair',
		]);
		const box = shapes[0]!.shape!;
		expect(box.geometry).toBe('rect');
		expect(box.fill).toMatchObject({ kind: 'solid', color: { value: 'FF8000' } });
		expect(box.line?.widthEmu).toBe(25400);
		expect(box.text?.text).toBe('Hello');
		expect(shapes[2]!.shape!.geometry).toBe('rightArrow');
		expect(shapes[3]!.shape!.textBox).toBe(true);
		expect(shapes[4]!.shape!.kind).toBe('connector');
		const group = shapes[5]!.shape!;
		expect(group.children?.map((c) => c.name)).toEqual(['G1', 'G2']);
		expect(group.childFrame?.width).toBe(1397000);
	});

	it('reads form controls with their state and captions', async () => {
		const wb = await loadXlsx(fixture());
		const controls = wb.sheets[0]!.drawings.filter(
			(d): d is FormControlObject => d.kind === 'formControl',
		);
		expect(controls.map((c) => c.controlType)).toEqual([
			'checkbox',
			'radio',
			'radio',
			'dropdown',
			'list',
			'button',
			'spin',
			'scroll',
			'groupBox',
		]);
		expect(controls[0]).toMatchObject({
			caption: 'Enable',
			linkedCell: '$C$1',
			checked: 'checked',
		});
		expect(controls[1]).toMatchObject({ caption: 'One', firstButton: true });
		expect(controls[3]).toMatchObject({ listRange: '$A$2:$A$4', selected: 2, dropStyle: 'combo' });
		expect(controls[6]).toMatchObject({ max: 10, value: 3 });
		expect(controls[7]).toMatchObject({ horizontal: true, value: 40, max: 100 });
		expect(controls[8]!.caption).toBe('Group');
	});

	it('keeps the form control parts when saving', async () => {
		const wb = await loadXlsx(fixture());
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		const names = Object.keys(zip.files);
		expect(names.filter((n) => n.includes('ctrlProp')).length).toBe(9);
		expect(names.some((n) => n.endsWith('.vml'))).toBe(true);
		const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(sheet).toContain('<control ');
		expect(sheet).toContain('legacyDrawing');
		const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
		expect(drawing.match(/compatExt/g)?.length).toBe(9);
	});

	it('keeps the controls and their VML when the drawing is regenerated', async () => {
		const wb = await loadXlsx(fixture());
		const box = wb.sheets[0]!.drawings.find(
			(d) => d.kind === 'unsupported' && d.shape?.name === 'Box',
		)!;
		box.anchor.from.col += 1;
		if (box.anchor.to) box.anchor.to.col += 1;
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
		expect(drawing.match(/compatExt/g)?.length).toBe(9);
		const vml = Object.keys(zip.files).find((n) => n.endsWith('.vml'))!;
		expect(await zip.file(vml)!.async('string')).toContain('ObjectType="Checkbox"');
		const again = await loadXlsx(
			await saveXlsx(await loadXlsx(await zip.generateAsync({ type: 'uint8array' }))),
		);
		expect(again.sheets[0]!.drawings.filter((d) => d.kind === 'formControl')).toHaveLength(9);
	});

	it('writes the state of the linked cells into the form control parts', async () => {
		const wb = await loadXlsx(fixture());
		const session = createEditSession(wb);
		session.setCellValue(0, 0, 2, false);
		session.setCellValue(0, 1, 2, 2);
		session.setCellValue(0, 5, 2, 80);
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		const part = async (n: number) => zip.file(`xl/ctrlProps/ctrlProp${n}.xml`)!.async('string');
		expect(await part(1)).not.toContain('checked=');
		expect(await part(8)).toContain('val="80"');
		expect(await part(2)).toContain('fmlaLink="$C$2"');
		expect(await part(2)).not.toContain('checked=');
		expect(await part(3)).toContain('checked="Checked"');
		expect(await part(6)).toBe(
			await JSZip.loadAsync(fixture()).then((z) =>
				z.file('xl/ctrlProps/ctrlProp6.xml')!.async('string'),
			),
		);
	});
});
