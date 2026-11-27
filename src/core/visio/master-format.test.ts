import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { readVisioStencil } from './stencil-file';
import { createVisioStencil, editVisioStencil } from './stencil-editor';
import { parseXml } from '../xml/index';
import {
	applyVisioMasterFormat,
	refreshInheritedCaches,
	snapshotVisioMasterFormat,
} from './master-format';

const NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const masterXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('visio/masters/master1.xml')!.async('string');

/** A drawing with one Rectangle instance dropped from the built-in Basic Shapes. */
async function drawing(): Promise<Uint8Array> {
	return (
		await editVsdx(await createVsdx(), [
			{ type: 'drop-stencil-master', pageId: '0', shapeId: '1', master: 'rectangle', x: 2, y: 2 },
		])
	).bytes;
}
const edit = (format: Parameters<typeof snapshotVisioMasterFormat>[0]) =>
	({ type: 'edit-document-master', pageId: '0', masterId: '1', format }) as const;

describe('Edit Master on a drawing master', () => {
	it('changes fill, line and text formatting and the instances show it', async () => {
		const source = await drawing();
		const result = await editVsdx(source, [
			edit({
				fillColor: '#ff0000',
				lineColor: '#00ff00',
				lineWeight: 2,
				fontColor: '#0000ff',
				fontSize: 18,
				bold: true,
				horizontalAlign: 'left',
			}),
		]);
		const xml = await masterXml(result.bytes);
		expect(xml).toContain('<Cell N="FillForegnd" V="#ff0000" F="RGB(255,0,0)"/>');
		expect(xml).toContain('<Cell N="LineColor" V="#00ff00" F="RGB(0,255,0)"/>');
		const shape = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!;
		expect(shape.style.fill).toBe('#ff0000');
		expect(shape.style.lineColor).toBe('#00ff00');
		expect(shape.style.lineWidth).toBeCloseTo(2 / 72);
		expect(shape.text.color).toBe('#0000ff');
		expect(shape.text.fontSize).toBeCloseTo(0.25);
		expect(shape.text.bold).toBe(true);
	});

	it('keeps formatting an instance set itself', async () => {
		const formatted = await editVsdx(await drawing(), [
			{ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#123456' },
		]);
		const result = await editVsdx(formatted.bytes, [edit({ fillColor: '#ff0000' })]);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.style.fill).toBe('#123456');
	});

	it('changes the default text, which instances without text of their own show', async () => {
		const result = await editVsdx(await drawing(), [edit({ text: 'Step' })]);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.text.plainText).toBe('Step');
	});

	it('refuses to resize a master whose instances take their size from it', async () => {
		await expect(editVsdx(await drawing(), [edit({ width: 3 })])).rejects.toThrow(
			/take their size from this master/,
		);
	});

	it('refuses a mixed transaction and invalid formats', async () => {
		const source = await drawing();
		await expect(
			editVsdx(source, [
				edit({ fillColor: '#ff0000' }),
				{ type: 'move-shape', pageId: '0', shapeId: '1', x: 3, y: 3 },
			]),
		).rejects.toThrow(/own transaction/);
		await expect(editVsdx(source, [edit({ fillColor: 'red' })])).rejects.toThrow(/#rrggbb/);
		await expect(editVsdx(source, [{ ...edit({ bold: true }), masterId: '9' }])).rejects.toThrow(
			/no such master/,
		);
	});
});

describe('Edit Master in a stencil', () => {
	it('resizes the master, recalculating the cells that follow its size', async () => {
		const stencil = (
			await editVisioStencil(await createVisioStencil(), {
				type: 'add-built-in',
				id: 'rectangle',
			})
		).bytes;
		const listed = await readVisioStencil(stencil);
		const master = listed.masters[0]!;
		const saved = await editVisioStencil(stencil, {
			type: 'edit-master-format',
			id: master.id,
			shapeId: master.shapes[0]!.id,
			format: { width: 2, height: 1.5, fillColor: '#00ff00', italic: true },
		});
		const xml = await masterXml(saved.bytes);
		expect(xml).toContain('<Cell N="Width" V="2" U="IN"/>');
		expect(xml).toContain('<Cell N="LocPinX" V="1" F="Width*0.5"/>');
		const after = (await readVisioStencil(saved.bytes)).masters[0]!.shapes[0]!;
		expect([after.width, after.height]).toEqual([2, 1.5]);
		expect(after.style.fill).toBe('#00ff00');
		expect(after.text.italic).toBe(true);
	});
});

describe('master format helpers', () => {
	const shape = () =>
		parseXml(
			`<Shape xmlns="${NS}" ID="5" Type="Shape"><Cell N="Width" V="1"/><Cell N="Height" V="1"/><Cell N="TxtWidth" V="1" F="Width*1"/><Cell N="Odd" V="1" F="Width+User.Pad"/></Shape>`,
		).documentElement;
	const document = parseXml(`<VisioDocument xmlns="${NS}"/>`).documentElement;

	it('refuses a size change a cell computes in a way it cannot follow', () => {
		expect(() => applyVisioMasterFormat(shape(), document, { width: 2 })).toThrow(
			/Width\+User\.Pad/,
		);
	});

	it('refreshes inherited caches and leaves local values', () => {
		const instance = parseXml(
			`<Shape xmlns="${NS}" ID="1" Master="1"><Cell N="FillForegnd" V="#000000" F="Inh"/><Cell N="LineColor" V="#111111"/></Shape>`,
		).documentElement;
		expect(
			refreshInheritedCaches(instance, [
				{ name: 'FillForegnd', value: '#ff0000' },
				{ name: 'LineColor', value: '#00ff00' },
			]),
		).toBe(true);
		expect(instance.getElementsByTagNameNS(NS, 'Cell')[0]!.getAttribute('V')).toBe('#ff0000');
		expect(instance.getElementsByTagNameNS(NS, 'Cell')[1]!.getAttribute('V')).toBe('#111111');
	});

	it('validates the format', () => {
		expect(snapshotVisioMasterFormat({ bold: true, text: 'a\r\nb' })).toEqual({
			bold: true,
			text: 'a\nb',
		});
		expect(() => snapshotVisioMasterFormat({ width: -1 })).toThrow();
		expect(() => snapshotVisioMasterFormat({ horizontalAlign: 'middle' as 'left' })).toThrow();
	});
});
