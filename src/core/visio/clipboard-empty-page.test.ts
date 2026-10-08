import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseXml } from '../xml/index';
import { captureVisioClipboard } from './clipboard';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { fixture, shape, cell, rectangle } from './test-fixtures';
import { visioPasteCommand } from './ui/shape-clipboard';
import { children, VISIO_NS, VISIO_LEGACY_NS } from './sheet';

const local = shape(
	'1',
	cell('PinX', 2) +
		cell('PinY', 3) +
		cell('Width', 2) +
		cell('Height', 1) +
		rectangle +
		'<Text>Captured</Text>',
);
const bytes = (contents: string, namespace: string) =>
	fixture({
		pages: [{ id: '0', contents }],
		edit: (zip) =>
			zip.file(
				'visio/pages/page1.xml',
				`<PageContents xmlns="${namespace}">${contents}</PageContents>`,
			),
	});

describe('paste onto empty source pages', () => {
	it.each([VISIO_NS, VISIO_LEGACY_NS])(
		'creates a namespace-correct Shapes container before Connects in %s',
		async (namespace) => {
			const snapshot = await captureVisioClipboard(
				await bytes(`<Shapes>${local}</Shapes>`, namespace),
				'0',
				['1'],
			);
			const target = await bytes('<Unknown Keep="yes"/><Connects/>', namespace);
			const before = target.slice();
			const page = (await parseVsdx(target)).pages[0]!;
			expect(page.shapes).toHaveLength(0);
			const command = visioPasteCommand(page, snapshot, { x: 0, y: 0 })!;
			expect(command.copies).toEqual([{ shapeId: '1', newShapeId: '1' }]);
			const saved = await editVsdx(target, [command]);
			const parsed = (await parseVsdx(saved.bytes)).pages[0]!;
			expect(parsed.shapes.map((shape) => shape.text.plainText)).toEqual(['Captured']);
			const root = parseXml(
				await (await JSZip.loadAsync(saved.bytes)).file('visio/pages/page1.xml')!.async('string'),
			).documentElement;
			expect(children(root, 'Shapes')).toHaveLength(1);
			expect(children(root, 'Shapes')[0]!.namespaceURI).toBe(namespace);
			expect(children(children(root, 'Shapes')[0], 'Shape')[0]!.namespaceURI).toBe(namespace);
			expect(
				Array.from(root.childNodes)
					.filter((node) => node.nodeType === 1)
					.map((node) => (node as Element).localName),
			).toEqual(['Unknown', 'Shapes', 'Connects']);
			expect(children(root, 'Unknown')[0]!.getAttribute('Keep')).toBe('yes');
			expect(target).toEqual(before);
		},
	);
	it('refuses ambiguous target containers atomically', async () => {
		const snapshot = await captureVisioClipboard(
			await bytes(`<Shapes>${local}</Shapes>`, VISIO_NS),
			'0',
			['1'],
		);
		const target = await bytes('<Shapes/><Shapes/>', VISIO_NS),
			before = target.slice();
		await expect(
			editVsdx(target, [
				{
					type: 'paste-shapes',
					pageId: '0',
					clipboard: snapshot,
					copies: [{ shapeId: '1', newShapeId: '2' }],
					offsetX: 0,
					offsetY: 0,
				},
			]),
		).rejects.toMatchObject({ code: 'INVALID_SHAPE_ID' });
		expect(target).toEqual(before);
	});
});
