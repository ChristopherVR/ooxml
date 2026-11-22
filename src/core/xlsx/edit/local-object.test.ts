import { describe, expect, it } from 'vitest';
import { embeddableOfficePackageFormat, openZip, RELATIONSHIP_TYPES } from '../../opc/index';
import { pictureAnchorAt } from '../layout/anchors';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { createEditSession } from './session';

const preview = Uint8Array.from(
	atob(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/2fkAAAAASUVORK5CYII=',
	),
	(c) => c.charCodeAt(0),
);

function modelBytes(): Uint8Array {
	const json = new TextEncoder().encode('{"asset":{"version":"2.0"}} ');
	const data = new Uint8Array(20 + json.length);
	const header = new DataView(data.buffer);
	header.setUint32(0, 0x46546c67, true);
	header.setUint32(4, 2, true);
	header.setUint32(8, data.length, true);
	header.setUint32(12, json.length, true);
	header.setUint32(16, 0x4e4f534a, true);
	data.set(json, 20);
	return data;
}

describe('local worksheet objects', () => {
	it('stores the editable model, embedded document and matching fallback pictures', async () => {
		const original = createWorkbook();
		const session = createEditSession(original);
		const model = modelBytes();
		const embedded = await saveXlsx(createWorkbook());
		const format = embeddableOfficePackageFormat('xlsx');
		if (!format) throw new Error('Missing XLSX embedding format');
		const modelAnchor = pictureAnchorAt({ row: 2, col: 1 }, 160, 160);
		const objectAnchor = pictureAnchorAt({ row: 10, col: 3 }, 96, 96);
		session.insertLocalObject(0, {
			kind: 'model3d',
			name: 'Model.glb',
			bytes: model,
			preview,
			anchor: modelAnchor,
		});
		session.insertLocalObject(0, {
			kind: 'object',
			name: 'Source.xlsx',
			bytes: embedded,
			preview,
			anchor: objectAnchor,
			format,
		});
		session.undo();
		expect(original.sheets[0]?.drawings.map((drawing) => drawing.kind)).toEqual(['model3d']);
		session.redo();
		const saved = await saveXlsx(original);
		const zip = await openZip(saved, { label: 'Workbook' });
		const contentTypes = await zip.file('[Content_Types].xml')?.async('string');
		expect(contentTypes).toContain('model/gltf.binary');
		expect(await zip.file('xl/drawings/drawing1.xml')?.async('string')).toContain('<am3d:model3d');
		expect(await zip.file('xl/worksheets/sheet1.xml')?.async('string')).toContain(
			'progId="Excel.Sheet.12"',
		);
		const reopened = await loadXlsx(saved);
		const [modelFrame, objectFrame] = reopened.sheets[0]?.drawings ?? [];
		expect(modelFrame).toMatchObject({ kind: 'model3d', name: 'Model.glb', anchor: modelAnchor });
		expect(objectFrame).toMatchObject({
			kind: 'embeddedObject',
			name: 'Source.xlsx',
			anchor: objectAnchor,
		});
		if (modelFrame?.kind !== 'model3d' || objectFrame?.kind !== 'embeddedObject')
			throw new Error('Local worksheet objects were not read');
		expect(reopened.source?.parts.get(modelFrame.partName)).toEqual(model);
		expect(reopened.source?.parts.get(objectFrame.partName)).toEqual(embedded);
		const second = createEditSession(reopened);
		second.setDrawingAnchor(0, 0, objectAnchor);
		second.deleteDrawing(0, 1);
		const modified = await loadXlsx(await saveXlsx(reopened));
		expect(modified.sheets[0]?.drawings).toHaveLength(1);
		expect(modified.sheets[0]?.drawings[0]?.anchor).toEqual(objectAnchor);
		expect(modified.source?.parts.has(objectFrame.partName)).toBe(false);
		expect(modified.source?.parts.has(objectFrame.previewPartName)).toBe(false);
	});

	it('preserves legacy worksheet objects alongside and after removing a local object', async () => {
		const workbook = createWorkbook();
		const format = embeddableOfficePackageFormat('xlsx');
		if (!format) throw new Error('Missing XLSX embedding format');
		createEditSession(workbook).insertLocalObject(0, {
			kind: 'object',
			name: 'Source.xlsx',
			bytes: await saveXlsx(createWorkbook()),
			format,
			preview,
			anchor: pictureAnchorAt({ row: 0, col: 0 }, 96, 96),
		});
		const sheetPart = 'xl/worksheets/sheet1.xml';
		const relPart = 'xl/worksheets/_rels/sheet1.xml.rels';
		const legacyPart = 'xl/embeddings/legacy.bin';
		const original = await openZip(await saveXlsx(workbook), { label: 'Workbook' });
		const sheetXml = (await original.file(sheetPart)?.async('string')) ?? '';
		const relsXml = (await original.file(relPart)?.async('string')) ?? '';
		const legacyObject =
			'<oleObject progId="Package" shapeId="1025" dvAspect="DVASPECT_ICON" r:id="rIdLegacy"/>';
		const legacyRel = `<Relationship Id="rIdLegacy" Type="${RELATIONSHIP_TYPES.oleObject}" Target="../embeddings/legacy.bin"/>`;
		original.file(sheetPart, sheetXml.replace('</oleObjects>', legacyObject + '</oleObjects>'));
		original.file(relPart, relsXml.replace('</Relationships>', legacyRel + '</Relationships>'));
		original.file(legacyPart, 'Legacy file payload');
		const mixed = await loadXlsx(await original.generateAsync({ type: 'uint8array' }));
		const written = await openZip(await saveXlsx(mixed), { label: 'Workbook' });
		expect(await written.file(sheetPart)?.async('string')).toContain('rIdLegacy');
		expect(await written.file(sheetPart)?.async('string')).toContain('Excel.Sheet.12');
		createEditSession(mixed).deleteDrawing(0, 0);
		const removed = await openZip(await saveXlsx(mixed), { label: 'Workbook' });
		const remaining = await removed.file(sheetPart)?.async('string');
		expect(remaining).toContain('rIdLegacy');
		expect(remaining).not.toContain('Excel.Sheet.12');
		expect(await removed.file(relPart)?.async('string')).toContain(legacyRel);
		expect(await removed.file(legacyPart)?.async('string')).toBe('Legacy file payload');
	});
});
