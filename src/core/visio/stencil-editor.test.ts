import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { decodeBase64 } from '../digest/base64';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { STENCIL_MACRO_MAIN, STENCIL_MAIN } from './package-kind';
import { VisioPackage } from './package';
import { children } from './sheet';
import { createVisioStencil, editVisioStencil, type VisioStencilEdit } from './stencil-editor';
import { readVisioStencil } from './stencil-file';
import { VISIO_BUILT_IN_STENCILS } from './stencil-masters';

const apply = async (bytes: Uint8Array, edit: VisioStencilEdit) =>
	(await editVisioStencil(bytes, edit)).bytes;
const builtIn = (bytes: Uint8Array, id: string) => apply(bytes, { type: 'add-built-in', id });
const masters = (bytes: Uint8Array) => readVisioStencil(bytes).then((file) => file.masters);

describe('stencil authoring', () => {
	it('creates a stencil window without a drawing page reference', async () => {
		const pkg = await VisioPackage.open(await createVisioStencil());
		const windows = children(await pkg.readXml('visio/windows.xml', 'Windows'), 'Window');
		expect(windows).toHaveLength(1);
		expect(windows[0]?.getAttribute('WindowType')).toBe('Stencil');
		expect(windows[0]?.hasAttribute('Page')).toBe(false);
		expect(windows[0]?.hasAttribute('ContainerType')).toBe(false);
	});
	it('chooses a name that differs from both local and universal names regardless of case', async () => {
		let stencil = await createVisioStencil();
		for (const id of ['triangle', 'ellipse', 'diamond']) stencil = await builtIn(stencil, id);
		const zip = await JSZip.loadAsync(stencil);
		const path = 'visio/masters/masters.xml';
		const xml = await zip.file(path)!.async('string');
		zip.file(
			path,
			xml
				.replace('Name="Triangle"', 'Name="circle"')
				.replace('NameU="Ellipse"', 'NameU="CIRCLE.4"')
				.replace('Name="Diamond"', 'Name="Circle.4.1"'),
		);
		const result = await builtIn(await zip.generateAsync({ type: 'uint8array' }), 'circle');
		expect((await masters(result)).at(-1)?.name).toBe('Circle.4.2');
	});
	it('writes every built-in as a ShapeSheet with geometry, protection, points, data and an icon', async () => {
		for (const stencil of VISIO_BUILT_IN_STENCILS) {
			let bytes = await createVisioStencil();
			expect(await masters(bytes)).toEqual([]);
			for (const item of stencil.masters) bytes = await builtIn(bytes, item.id);
			const read = await masters(bytes);
			expect(read.map((item) => item.name)).toEqual(stencil.masters.map((item) => item.name));
			for (const [index, master] of read.entries()) {
				const item = stencil.masters[index]!;
				expect(master.shapes[0]).toMatchObject({
					...item.size,
					placeable: true,
					shapeData: [{ name: 'ShapeKind', label: 'Shape type', value: item.name }],
				});
				expect(master.shapes[0]?.connectionPoints).toHaveLength(4);
				expect(master.shapes[0]?.geometry.length).toBeGreaterThan(0);
			}
			const zip = await JSZip.loadAsync(bytes);
			const xml = await zip.file('visio/masters/masters.xml')!.async('string');
			const icons = Array.from(xml.matchAll(/<Icon>([^<]+)<\/Icon>/g), ([, data]) =>
				decodeBase64(data ?? ''),
			);
			expect(icons).toHaveLength(read.length);
			for (const icon of icons) {
				expect(icon?.slice(0, 9)).toEqual(new Uint8Array([0, 0, 1, 0, 1, 0, 32, 32, 16]));
			}
		}
	}, 90_000);

	it('adds and edits a drawing shape without randomUUID, then drops and deletes it', async () => {
		vi.stubGlobal('crypto', { getRandomValues: crypto.getRandomValues.bind(crypto) });
		try {
			await authorDrawingShape();
		} finally {
			vi.unstubAllGlobals();
		}
	});

	async function authorDrawingShape() {
		const drawing = (
			await editVsdx(await createVsdx(), [
				{
					type: 'create-rectangle',
					pageId: '0',
					shapeId: '1',
					x: 4,
					y: 5,
					width: 2,
					height: 1,
					text: 'Source',
				},
				{ type: 'add-connection-point', pageId: '0', shapeId: '1', x: 0, y: 0.5 },
			])
		).bytes;
		const original = await apply(await createVisioStencil(), {
			type: 'add-shape',
			drawing,
			pageId: '0',
			shapeId: '1',
			name: 'Valve',
		});
		const master = (await masters(original))[0]!;
		expect(master.shapes[0]).toMatchObject({ width: 2, height: 1, text: { plainText: 'Source' } });
		let stencil = await apply(original, {
			type: 'rename-master',
			id: master.id,
			name: 'Spare <valve>',
		});
		stencil = await apply(stencil, {
			type: 'edit-master-text',
			id: master.id,
			shapeId: '1',
			text: 'A & B\nSecond line',
		});
		const [edited] = await masters(stencil);
		expect(edited).toMatchObject({ name: 'Spare <valve>' });
		const result = await editVsdx(await createVsdx(), [
			{
				type: 'drop-stencil-master',
				pageId: '0',
				shapeId: '2',
				master: master.id,
				stencil,
				x: 3,
				y: 4,
			},
		]);
		expect((await parseVsdx(result.bytes)).pages[0]?.shapes[0]).toMatchObject({
			masterId: '1',
			width: 2,
			height: 1,
			text: { plainText: 'A & B\nSecond line' },
			connectionPoints: [{ x: 0, y: 0.5, inherited: true }],
		});
		const removed = await apply(stencil, { type: 'delete-master', id: master.id });
		expect(await masters(removed)).toEqual([]);
		expect((await JSZip.loadAsync(removed)).file('visio/masters/master1.xml')).toBeNull();
		expect((await masters(original))[0]?.name).toBe('Valve');
	}

	it('preserves a macro stencil project and settings through edits without executing code', async () => {
		const source = await builtIn(await createVisioStencil(), 'circle');
		const zip = await JSZip.loadAsync(source);
		const types = await zip.file('[Content_Types].xml')!.async('string');
		zip.file('[Content_Types].xml', types.replace(STENCIL_MAIN, STENCIL_MACRO_MAIN));
		const relsPath = 'visio/_rels/document.xml.rels';
		const rels = await zip.file(relsPath)!.async('string');
		zip.file(
			relsPath,
			rels.replace(
				'</Relationships>',
				'<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
			),
		);
		const vba = Uint8Array.of(19, 22, 255, 12);
		zip.file('visio/vbaProject.bin', vba);
		const loaded = await zip.generateAsync({ type: 'uint8array' });
		const result = await editVisioStencil(loaded, {
			type: 'rename-master',
			id: '1',
			name: 'Circle A',
		});
		expect(result.extension).toBe('vssm');
		expect((await masters(result.bytes))[0]?.name).toBe('Circle A');
		const reopened = await JSZip.loadAsync(result.bytes);
		expect(await reopened.file('visio/vbaProject.bin')?.async('uint8array')).toEqual(vba);
		expect(await reopened.file(relsPath)?.async('string')).toContain('rIdVba');
		expect(await reopened.file('visio/document.xml')?.async('string')).toBe(
			await zip.file('visio/document.xml')?.async('string'),
		);
		const copied = await apply(await createVisioStencil(), {
			type: 'import-master',
			stencil: result.bytes,
			id: '1',
		});
		expect((await masters(copied))[0]?.name).toBe('Circle A');
		expect((await JSZip.loadAsync(copied)).file('visio/vbaProject.bin')).toBeNull();
	});

	it('removes unreferenced master media and its content types, retaining shared and opaque parts', async () => {
		const source = await builtIn(await builtIn(await createVisioStencil(), 'rectangle'), 'circle');
		const zip = await JSZip.loadAsync(source);
		const typesPath = '[Content_Types].xml';
		const types = await zip.file(typesPath)!.async('string');
		zip.file(
			typesPath,
			types.replace(
				'</Types>',
				'<Default Extension="png" ContentType="image/png"/><Default Extension="gif" ContentType="image/gif"/>' +
					'<Override PartName="/visio/media/only.gif" ContentType="image/gif"/></Types>',
			),
		);
		const rels = (targets: string[]) =>
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			targets
				.map(
					(target, index) =>
						`<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${target}"/>`,
				)
				.join('') +
			'</Relationships>';
		const rels1 = 'visio/masters/_rels/master1.xml.rels';
		zip.file(rels1, rels(['shared.png', 'only.gif']));
		zip.file('visio/masters/_rels/master2.xml.rels', rels(['shared.png']));
		for (const part of ['shared.png', 'only.gif', 'opaque.png'])
			zip.file(`visio/media/${part}`, part);
		const edited = await apply(await zip.generateAsync({ type: 'uint8array' }), {
			type: 'delete-master',
			id: '1',
		});
		const result = await JSZip.loadAsync(edited);
		expect((await masters(edited)).map(({ name }) => name)).toEqual(['Circle']);
		expect(result.file(rels1)).toBeNull();
		expect(result.file('visio/media/only.gif')).toBeNull();
		expect(await result.file('visio/media/shared.png')?.async('string')).toBe('shared.png');
		expect(await result.file('visio/media/opaque.png')?.async('string')).toBe('opaque.png');
		const savedTypes = await result.file(typesPath)?.async('string');
		expect(savedTypes).not.toContain('gif');
		expect(savedTypes).toContain('Extension="png"');
	});

	it('rejects oversize inputs and hostile XML before publishing a change', async () => {
		const source = await builtIn(await createVisioStencil(), 'rectangle');
		const rename = { type: 'rename-master', id: '1', name: 'Box' } satisfies VisioStencilEdit;
		await expect(
			editVisioStencil(source, rename, { limits: { maxInputBytes: 128 } }),
		).rejects.toMatchObject({ code: 'LIMIT_INPUT' });
		await expect(
			editVisioStencil(source, rename, { limits: { maxEntries: 3 } }),
		).rejects.toMatchObject({ code: 'LIMIT_ENTRIES' });
		const zip = await JSZip.loadAsync(source);
		const part = 'visio/masters/masters.xml';
		const xml = await zip.file(part)!.async('string');
		zip.file(part, '<!DOCTYPE Masters>' + xml);
		const hostile = await zip.generateAsync({ type: 'uint8array' });
		await expect(editVisioStencil(hostile, rename)).rejects.toThrow(/DOCTYPE|DTD/i);
		await expect(readVisioStencil(hostile)).rejects.toThrow(/DOCTYPE|DTD/i);
		expect((await masters(source))[0]?.name).toBe('Rectangle');
	});
});
