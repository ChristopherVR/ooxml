import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { PresentationBuilder } from './PresentationBuilder';
import { createSlideClipboardPackage, importSlidePackage } from './import-slide-package';
import { PptxHandler } from '../../PptxHandler';
import { NS, parseXml } from '../../../../xml/index';
import {
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
} from '../../../../opc/relationships';

async function deck(label: string, color: string, notes?: string) {
	const { handler, data, createSlide } = await PresentationBuilder.create({
		theme: { colors: { accent1: color } },
	});
	data.slides = [createSlide('Blank').addText(label).build()];
	if (notes) {
		data.slides[0]!.notes = notes;
		data.slides[0]!.notesSegments = [{ text: notes, style: { bold: true } }];
	}
	return handler.save(data.slides);
}
const arrayBuffer = (bytes: Uint8Array) =>
	bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

it('allocates master and layout native IDs independently across repeated imports', async () => {
	const destination = await deck('Destination', '#112233');
	const source = await deck('Source', '#ABCDEF');
	const first = await importSlidePackage(destination, source);
	const second = await importSlidePackage(first.bytes, source);
	const zip = await JSZip.loadAsync(second.bytes);
	const ids: number[] = [];
	const presentation = parseXml(await zip.file('ppt/presentation.xml')!.async('string'));
	for (const entry of Array.from(presentation.getElementsByTagNameNS(NS.p, 'sldMasterId')))
		ids.push(Number(entry.getAttribute('id')));
	for (const path of Object.keys(zip.files)) {
		if (!/^ppt\/slideMasters\/[^/]+\.xml$/u.test(path)) continue;
		const master = parseXml(await zip.file(path)!.async('string'));
		for (const entry of Array.from(master.getElementsByTagNameNS(NS.p, 'sldLayoutId')))
			ids.push(Number(entry.getAttribute('id')));
	}
	expect(new Set(ids).size).toBe(ids.length);
	expect(ids.every((id) => id >= 2147483648)).toBe(true);
});

describe('cross-deck package slide import', () => {
	it('retains independent master themes, notes, native IDs and opaque source XML after save/reopen', async () => {
		const destination = await deck('Destination', '#112233');
		const source = await deck('Imported', '#ABCDEF', 'Source speaker notes');
		const sourceZip = await JSZip.loadAsync(source);
		const sourceSlide = 'ppt/slides/slide1.xml';
		sourceZip.file(
			sourceSlide,
			(await sourceZip.file(sourceSlide)!.async('string')).replace(
				'</p:sld>',
				'<p:extLst><p:ext uri="opaque-import"/></p:extLst></p:sld>',
			),
		);
		const result = await importSlidePackage(
			destination,
			await sourceZip.generateAsync({ type: 'uint8array' }),
			{ insertAt: 0 },
		);
		const importedZip = await JSZip.loadAsync(result.bytes);
		expect(result.slidePaths).toEqual(['ppt/slides/slide2.xml']);
		expect(await importedZip.file('ppt/theme/theme1.xml')!.async('string')).toBe(
			await (await JSZip.loadAsync(destination)).file('ppt/theme/theme1.xml')!.async('string'),
		);
		expect(await importedZip.file('ppt/theme/theme2.xml')!.async('string')).toContain('ABCDEF');
		expect(await importedZip.file(result.slidePaths[0]!)!.async('string')).toContain(
			'opaque-import',
		);
		const handler = new PptxHandler(),
			loaded = await handler.load(arrayBuffer(result.bytes));
		expect(
			loaded.slides.map((slide) => slide.elements.find((el) => el.type === 'text')?.text),
		).toEqual(['Imported', 'Destination']);
		expect(loaded.slides[0]!.notes).toContain('Source speaker notes');
		expect(
			loaded.slides[0]!.notesSegments?.find((run) => run.text.includes('Source'))?.style.bold,
		).toBe(true);
		expect(loaded.slideMasters).toHaveLength(2);
		const saved = await handler.save(loaded.slides),
			reopened = await new PptxHandler().load(arrayBuffer(saved));
		expect(reopened.slides).toHaveLength(2);
		expect(reopened.slides[0]!.notes).toContain('Source speaker notes');
		const savedZip = await JSZip.loadAsync(saved);
		expect(await savedZip.file(result.slidePaths[0]!)!.async('string')).toContain('opaque-import');
	});
	it('copies media and embedded dependency bytes despite source/destination path collisions', async () => {
		const source = await JSZip.loadAsync(await deck('Imported', '#ABCDEF'));
		const destination = await JSZip.loadAsync(await deck('Destination', '#112233'));
		const path = 'ppt/slides/slide1.xml';
		const sourceRels = parseRelationships(
			await source.file(relationshipsPartFor(path))!.async('string'),
		);
		sourceRels.set('customMediaId', {
			id: 'customMediaId',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
			target: '../media/image1.png',
			mode: 'Internal',
		});
		sourceRels.set('nativeExternal', {
			id: 'nativeExternal',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
			target: 'https://example.com/',
			mode: 'External',
		});
		const { buildRelationshipsXml } = await import('../../../../opc/relationships');
		source.file(relationshipsPartFor(path), buildRelationshipsXml(sourceRels));
		source.file('ppt/media/image1.png', new Uint8Array([1, 3, 5]));
		destination.file('ppt/media/image1.png', new Uint8Array([2, 4, 6]));
		source.file(
			'[Content_Types].xml',
			(await source.file('[Content_Types].xml')!.async('string')).replace(
				'</Types>',
				'<Default Extension="png" ContentType="image/png"/></Types>',
			),
		);
		const result = await importSlidePackage(
			await destination.generateAsync({ type: 'uint8array' }),
			await source.generateAsync({ type: 'uint8array' }),
		);
		const output = await JSZip.loadAsync(result.bytes);
		const rels = parseRelationships(
			await output.file(relationshipsPartFor(result.slidePaths[0]!))!.async('string'),
		);
		const mediaPath = resolvePartPath(result.slidePaths[0]!, rels.get('customMediaId')!.target);
		expect(mediaPath).not.toBe('ppt/media/image1.png');
		expect(await output.file(mediaPath)!.async('uint8array')).toEqual(new Uint8Array([1, 3, 5]));
		expect(await output.file('ppt/media/image1.png')!.async('uint8array')).toEqual(
			new Uint8Array([2, 4, 6]),
		);
		expect(rels.get('nativeExternal')).toMatchObject({
			target: 'https://example.com/',
			mode: 'External',
		});
	});
});

it('rejects unsupported globals, missing source parts and links outside the selected set explicitly', async () => {
	const sourceBytes = await deck('Source', '#ABCDEF'),
		destination = await deck('Destination', '#112233');
	await expect(importSlidePackage(destination, sourceBytes, { slideIndices: [5] })).rejects.toThrow(
		'indexes',
	);
	await expect(
		importSlidePackage(destination, sourceBytes, { insertAt: Number.NaN }),
	).rejects.toThrow('finite integer');
	const source = await JSZip.loadAsync(sourceBytes);
	const { buildRelationshipsXml } = await import('../../../../opc/relationships');
	const relsPath = relationshipsPartFor('ppt/slides/slide1.xml');
	const relationships = parseRelationships(await source.file(relsPath)!.async('string'));
	relationships.set('missingImage', {
		id: 'missingImage',
		type: 'image',
		target: '../media/missing.png',
		mode: 'Internal',
	});
	source.file(relsPath, buildRelationshipsXml(relationships));
	await expect(
		importSlidePackage(destination, await source.generateAsync({ type: 'uint8array' })),
	).rejects.toThrow('missing.png');
	relationships.delete('missingImage');
	relationships.set('jump', {
		id: 'jump',
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
		target: 'slide9.xml',
		mode: 'Internal',
	});
	source.file(relsPath, buildRelationshipsXml(relationships));
	await expect(
		importSlidePackage(destination, await source.generateAsync({ type: 'uint8array' })),
	).rejects.toThrow('outside the selected set');
	relationships.delete('jump');
	relationships.set('comments', {
		id: 'comments',
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
		target: '../comments/comment1.xml',
		mode: 'Internal',
	});
	source.file(relsPath, buildRelationshipsXml(relationships));
	await expect(
		importSlidePackage(destination, await source.generateAsync({ type: 'uint8array' })),
	).rejects.toThrow('comments are not supported');
	source.file(
		'ppt/presentation.xml',
		(await source.file('ppt/presentation.xml')!.async('string')).replace(
			'</p:presentation>',
			'<p:embeddedFontLst/></p:presentation>',
		),
	);
	await expect(
		importSlidePackage(destination, await source.generateAsync({ type: 'uint8array' })),
	).rejects.toThrow('embedded fonts');
});

it('preserves selected slide order and remaps internal links among selected slides', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [
		createSlide('Blank').addText('First').build(),
		createSlide('Blank').addText('Second').build(),
	];
	const source = await JSZip.loadAsync(await handler.save(data.slides));
	const relsPath = relationshipsPartFor('ppt/slides/slide1.xml');
	const relationships = parseRelationships(await source.file(relsPath)!.async('string'));
	relationships.set('nativeJump', {
		id: 'nativeJump',
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
		target: 'slide2.xml',
		mode: 'Internal',
	});
	const { buildRelationshipsXml } = await import('../../../../opc/relationships');
	source.file(relsPath, buildRelationshipsXml(relationships));
	const destination = await deck('Destination', '#112233');
	const result = await importSlidePackage(
		destination,
		await source.generateAsync({ type: 'uint8array' }),
		{ slideIndices: [1, 0], insertAt: 0 },
	);
	const output = await JSZip.loadAsync(result.bytes);
	const rels = parseRelationships(
		await output.file(relationshipsPartFor(result.slidePaths[1]!))!.async('string'),
	);
	expect(resolvePartPath(result.slidePaths[1]!, rels.get('nativeJump')!.target)).toBe(
		result.slidePaths[0],
	);
	const reopened = await new PptxHandler().load(arrayBuffer(result.bytes));
	expect(
		reopened.slides.map((slide) => slide.elements.find((el) => el.type === 'text')?.text),
	).toEqual(['Second', 'First', 'Destination']);
	expect(reopened.slideMasters).toHaveLength(2);
});

it('captures only selected slides, excludes unselected content and is portable into another deck', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [
		createSlide('Blank').addText('Private unselected content').build(),
		createSlide('Blank').addText('Copied slide').build(),
	];
	const snapshot = await createSlideClipboardPackage(await handler.save(data.slides), [1]);
	const snapshotZip = await JSZip.loadAsync(snapshot);
	const xmlParts = await Promise.all(
		Object.values(snapshotZip.files)
			.filter((file) => !file.dir && file.name.endsWith('.xml'))
			.map((file) => file.async('string')),
	);
	expect(xmlParts.join('')).not.toContain('Private unselected content');
	const clipboardData = await new PptxHandler().load(arrayBuffer(snapshot));
	expect(clipboardData.slides).toHaveLength(1);
	expect(clipboardData.slides[0]!.elements.find((el) => el.type === 'text')?.text).toBe(
		'Copied slide',
	);
	const pasted = await importSlidePackage(await deck('Destination', '#112233'), snapshot);
	expect((await new PptxHandler().load(arrayBuffer(pasted.bytes))).slides).toHaveLength(2);
});
