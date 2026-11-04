import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { mergePptxSnapshots } from '../merge-snapshots';
import { NS } from '../../xml/namespaces';
import type { SnapshotMergeResult } from '../../opc/merge-snapshots';
import {
	createPresentationFromOutline,
	inspectPresentationText,
	setPresentationRunText,
} from '../core/utils/suite-text';

const shape = (id: number, text: string, x = 100) =>
	`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="100"/><a:ext cx="200" cy="100"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const slide = (shapes: string) =>
	`<p:sld xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`;
async function pptx(shapes: string, extra: Record<string, string | Uint8Array> = {}) {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="${NS.contentTypes}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
	);
	zip.file(
		'ppt/presentation.xml',
		`<p:presentation xmlns:p="${NS.p}" xmlns:r="${NS.r}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
	);
	zip.file(
		'ppt/_rels/presentation.xml.rels',
		`<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/slide" Target="slides/slide1.xml"/></Relationships>`,
	);
	zip.file('ppt/slides/slide1.xml', slide(shapes));
	for (const [name, bytes] of Object.entries(extra)) zip.file(name, bytes);
	return zip.generateAsync({ type: 'uint8array' });
}
async function mergedZip(result: SnapshotMergeResult) {
	expect(result.status, result.status === 'unresolved' ? result.issues.join('\n') : '').toBe(
		'merged',
	);
	if (result.status !== 'merged') throw new Error('Expected merged PPTX');
	return JSZip.loadAsync(result.bytes);
}
async function mergedXml(base: Uint8Array, local: Uint8Array, remote: Uint8Array) {
	return (await mergedZip(await mergePptxSnapshots(base, local, remote)))
		.file('ppt/slides/slide1.xml')!
		.async('string');
}

describe('mergePptxSnapshots', () => {
	it('keeps simultaneous typing at the same text position on both sides', async () => {
		const base = await pptx(shape(2, 'Hello'));
		const local = await pptx(shape(2, 'AHello'));
		const remote = await pptx(shape(2, 'BHello'));
		expect(await mergedXml(base, local, remote)).toContain('ABHello');
		expect(await mergedXml(base, remote, local)).toContain('ABHello');
	});

	it('merges two independent edits to one existing text run', async () => {
		const base = await pptx(shape(2, 'Hello world'));
		const xml = await mergedXml(
			base,
			await pptx(shape(2, 'Dear Hello world')),
			await pptx(shape(2, 'Hello world!')),
		);
		expect(xml).toContain('Dear Hello world!');
	});

	it('combines a text edit and geometry change on the same stable shape', async () => {
		const base = await pptx(shape(2, 'Title'));
		const xml = await mergedXml(
			base,
			await pptx(shape(2, 'New Title')),
			await pptx(shape(2, 'Title', 250)),
		);
		expect(xml).toContain('New Title');
		expect(xml).toContain('x="250"');
	});

	it('combines different shape edits without changing shape identities or z-order', async () => {
		const base = await pptx(shape(2, 'First') + shape(3, 'Second'));
		const xml = await mergedXml(
			base,
			await pptx(shape(2, 'Local First') + shape(3, 'Second')),
			await pptx(shape(2, 'First') + shape(3, 'Remote Second')),
		);
		expect(xml).toContain('Local First');
		expect(xml).toContain('Remote Second');
		expect(xml.indexOf('id="2"')).toBeLessThan(xml.indexOf('id="3"'));
	});

	it('combines independent position and fill changes while refusing the same position field', async () => {
		const base = await pptx(shape(2, 'Base'));
		const fill = shape(2, 'Base').replace(
			'</p:spPr>',
			'<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>',
		);
		const xml = await mergedXml(base, await pptx(shape(2, 'Base', 300)), await pptx(fill));
		expect(xml).toContain('x="300"');
		expect(xml).toContain('val="FF0000"');
		expect(
			(
				await mergePptxSnapshots(
					base,
					await pptx(shape(2, 'Base', 300)),
					await pptx(shape(2, 'Base', 400)),
				)
			).status,
		).toBe('unresolved');
	});

	it('retains new identified shapes and independently changed opaque media', async () => {
		const old = new Uint8Array([0]);
		const media = new Uint8Array([5, 4, 3, 2]);
		const base = await pptx(shape(2, 'Base'), { 'ppt/media/opaque.bin': old });
		const result = await mergePptxSnapshots(
			base,
			await pptx(shape(2, 'Base') + shape(3, 'New shape'), { 'ppt/media/opaque.bin': media }),
			await pptx(shape(2, 'Base remote'), { 'ppt/media/opaque.bin': old }),
		);
		const zip = await mergedZip(result);
		expect(await zip.file('ppt/media/opaque.bin')!.async('uint8array')).toEqual(media);
		const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(xml).toContain('New shape');
		expect(xml).toContain('Base remote');
	});

	it('retains unknown slide extensions while merging text and geometry', async () => {
		const decorate = (text: string, x = 100) =>
			shape(2, text, x).replace(
				'</p:sp>',
				'<p:extLst><p:ext uri="urn:merge-test"><custom:payload xmlns:custom="urn:merge-test" value="opaque">untouched</custom:payload></p:ext></p:extLst></p:sp>',
			);
		const xml = await mergedXml(
			await pptx(decorate('Base')),
			await pptx(decorate('Local Base')),
			await pptx(decorate('Base', 300)),
		);
		expect(xml).toContain('Local Base');
		expect(xml).toContain('x="300"');
		expect(xml).toContain('value="opaque"');
		expect(xml).toContain('untouched');
	});

	it('refuses overlapping text edits and deleting a shape the other side edits', async () => {
		const base = await pptx(shape(2, 'Original'));
		expect(
			(
				await mergePptxSnapshots(
					base,
					await pptx(shape(2, 'Local')),
					await pptx(shape(2, 'Remote')),
				)
			).status,
		).toBe('unresolved');
		expect(
			(await mergePptxSnapshots(base, await pptx(''), await pptx(shape(2, 'Remote')))).status,
		).toBe('unresolved');
	});

	it('refuses colliding new shape IDs and ambiguous run splits', async () => {
		const base = await pptx(shape(2, 'Base'));
		expect(
			(
				await mergePptxSnapshots(
					base,
					await pptx(shape(2, 'Base') + shape(3, 'Local added')),
					await pptx(shape(2, 'Base') + shape(3, 'Remote added')),
				)
			).status,
		).toBe('unresolved');
		const split = shape(2, 'Base').replace(
			'<a:t>Base</a:t>',
			'<a:t>Ba</a:t></a:r><a:r><a:t>se</a:t>',
		);
		expect(
			(await mergePptxSnapshots(base, await pptx(split), await pptx(shape(2, 'Base remote'))))
				.status,
		).toBe('unresolved');
	});

	it('keeps conflicting relationship IDs and binary parts unresolved', async () => {
		const rels = (target: string) =>
			`<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/hyperlink" Target="${target}" TargetMode="External"/></Relationships>`;
		const base = await pptx(shape(2, 'Base'), {
			'ppt/slides/_rels/slide1.xml.rels': rels('https://base.example'),
			'ppt/media/opaque.bin': new Uint8Array([0]),
		});
		const local = await pptx(shape(2, 'Base'), {
			'ppt/slides/_rels/slide1.xml.rels': rels('https://local.example'),
			'ppt/media/opaque.bin': new Uint8Array([1]),
		});
		const remote = await pptx(shape(2, 'Base'), {
			'ppt/slides/_rels/slide1.xml.rels': rels('https://remote.example'),
			'ppt/media/opaque.bin': new Uint8Array([2]),
		});
		expect((await mergePptxSnapshots(base, local, remote)).status).toBe('unresolved');
	});

	it('merges actual presentation-writer output and reopens the resulting deck', async () => {
		const base = await createPresentationFromOutline([{ title: 'Hello world' }]);
		const inspected = await inspectPresentationText(base);
		const element = inspected.slides[0]?.elements[0];
		if (!element) throw new Error('Expected generated title element');
		const local = await setPresentationRunText(base, 0, element.id, 0, 'Dear Hello world');
		const remote = await setPresentationRunText(base, 0, element.id, 0, 'Hello world!');
		const result = await mergePptxSnapshots(base, local, remote);
		await mergedZip(result);
		if (result.status === 'merged') {
			const reopened = await inspectPresentationText(result.bytes);
			expect(reopened.slides[0]?.elements.some((item) => item.text === 'Dear Hello world!')).toBe(
				true,
			);
		}
	}, 30_000);
});
