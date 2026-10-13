import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { PresentationBuilder } from './PresentationBuilder';
import { importSlidePackage } from './import-slide-package';
import { buildXml, NS, parseXml, type XmlElement } from '../../../../xml/index';
import { parseRelationships, resolvePartPath } from '../../../../opc/relationships';
import { renameFixtureParts } from '../../../__tests__/helpers/rename-package-parts';

async function deck(notes = false): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [createSlide('Blank').addText('ID regression').build()];
	if (notes) data.slides[0]!.notes = 'Notes';
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

it('reuses free valid IDs when destination slide, master and layout IDs are at their upper bounds', async () => {
	const zip = await JSZip.loadAsync(await deck());
	const presentation = parseXml(await zip.file('ppt/presentation.xml')!.async('string'));
	presentation.getElementsByTagNameNS(NS.p, 'sldId')[0]!.setAttribute('id', '2147483647');
	presentation.getElementsByTagNameNS(NS.p, 'sldMasterId')[0]!.setAttribute('id', '4294967295');
	zip.file('ppt/presentation.xml', buildXml(presentation));
	const master = parseXml(await zip.file('ppt/slideMasters/slideMaster1.xml')!.async('string'));
	master.getElementsByTagNameNS(NS.p, 'sldLayoutId')[0]!.setAttribute('id', '4294967294');
	zip.file('ppt/slideMasters/slideMaster1.xml', buildXml(master));
	const imported = await importSlidePackage(
		await zip.generateAsync({ type: 'uint8array' }),
		await deck(),
	);
	const output = await JSZip.loadAsync(imported.bytes);
	const result = parseXml(await output.file('ppt/presentation.xml')!.async('string'));
	expect(
		Array.from(result.getElementsByTagNameNS(NS.p, 'sldId')).map((node) =>
			Number(node.getAttribute('id')),
		),
	).toEqual([2147483647, 256]);
	const ids = Array.from(result.getElementsByTagNameNS(NS.p, 'sldMasterId')).map((node) =>
		Number(node.getAttribute('id')),
	);
	for (const relationship of parseRelationships(
		await output.file('ppt/_rels/presentation.xml.rels')!.async('string'),
	).values()) {
		if (!relationship.type.endsWith('/slideMaster')) continue;
		const doc = parseXml(
			await output
				.file(resolvePartPath('ppt/presentation.xml', relationship.target))!
				.async('string'),
		);
		ids.push(
			...Array.from(doc.getElementsByTagNameNS(NS.p, 'sldLayoutId')).map((node) =>
				Number(node.getAttribute('id')),
			),
		);
	}
	expect(new Set(ids).size).toBe(ids.length);
	expect(ids.every((id) => id >= 2147483648 && id <= 4294967295)).toBe(true);
});

it.each(['255', '2147483648', '4294967296'])(
	'rejects destination slide ID %s outside the slide format range',
	async (id) => {
		const zip = await JSZip.loadAsync(await deck());
		const doc = parseXml(await zip.file('ppt/presentation.xml')!.async('string'));
		doc.getElementsByTagNameNS(NS.p, 'sldId')[0]!.setAttribute('id', id);
		zip.file('ppt/presentation.xml', buildXml(doc));
		await expect(
			importSlidePackage(await zip.generateAsync({ type: 'uint8array' }), await deck()),
		).rejects.toThrow('Destination slide IDs');
	},
);

it('registers source master parts by relationship type with arbitrary valid names and directories', async () => {
	const source = await renameFixtureParts(
		await deck(true),
		new Map([
			['ppt/slideMasters/slideMaster1.xml', 'ppt/templates/company-design.xml'],
			['ppt/notesMasters/notesMaster1.xml', 'ppt/templates/presenter-notes.xml'],
		]),
	);
	const imported = await importSlidePackage(await deck(), source);
	const zip = await JSZip.loadAsync(imported.bytes);
	const presentation = parseXml(await zip.file('ppt/presentation.xml')!.async('string'));
	const relationships = parseRelationships(
		await zip.file('ppt/_rels/presentation.xml.rels')!.async('string'),
	);
	const masters = Array.from(presentation.getElementsByTagNameNS(NS.p, 'sldMasterId'));
	expect(masters).toHaveLength(2);
	const resolve = (node: XmlElement) =>
		resolvePartPath(
			'ppt/presentation.xml',
			relationships.get(node.getAttributeNS(NS.r, 'id')!)!.target,
		);
	expect(resolve(masters[1]!)).toContain('company-design');
	const notes = presentation.getElementsByTagNameNS(NS.p, 'notesMasterId')[0]!;
	expect(resolve(notes)).toContain('presenter-notes');
	const layoutIds = Array.from(
		parseXml(await zip.file(resolve(masters[1]!))!.async('string')).getElementsByTagNameNS(
			NS.p,
			'sldLayoutId',
		),
	).map((node) => Number(node.getAttribute('id')));
	expect(layoutIds.every((id) => id >= 2147483648 && id <= 4294967295)).toBe(true);
	// A second import must also inspect the destination's nonstandard master layout IDs.
	await expect(importSlidePackage(imported.bytes, source)).resolves.toBeDefined();
});

it.each(['sldMasterId', 'sldLayoutId'])(
	'rejects %s IDs above the unsigned 32-bit range',
	async (name) => {
		const zip = await JSZip.loadAsync(await deck());
		const path =
			name === 'sldMasterId' ? 'ppt/presentation.xml' : 'ppt/slideMasters/slideMaster1.xml';
		const doc = parseXml(await zip.file(path)!.async('string'));
		doc.getElementsByTagNameNS(NS.p, name)[0]!.setAttribute('id', '4294967296');
		zip.file(path, buildXml(doc));
		await expect(
			importSlidePackage(await zip.generateAsync({ type: 'uint8array' }), await deck()),
		).rejects.toThrow('Destination master/layout IDs');
	},
);
