import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	DIGITAL_SIGNATURE_ORIGIN_REL_TYPE,
	DIGITAL_SIGNATURE_REL_TYPE,
} from '../opc/signature/index';
import { createDocument } from './model';
import { loadDocx } from './parse';
import { saveDocx } from './save';
import { expectParagraph, must } from './test-support/access';

const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const SIGNATURE_TYPE = 'application/vnd.openxmlformats-package.digital-signature-';

// Synthetic OPC metadata, deliberately not a cryptographically valid signature.
async function signedDocument(folder = '_xmlsignatures'): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await saveDocx(createDocument()));
	zip.file(
		'_rels/.rels',
		`<r:Relationships xmlns:r="${RELS}"><r:Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><r:Relationship Id="sig" Type="${DIGITAL_SIGNATURE_ORIGIN_REL_TYPE}" Target="/${folder}/origin.sigs"/></r:Relationships>`,
	);
	zip.file(
		'[Content_Types].xml',
		`<ct:Types xmlns:ct="${TYPES}"><ct:Default Extension="xml" ContentType="application/xml"/><ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><ct:Default Extension="sigs" ContentType="${SIGNATURE_TYPE}origin"/><ct:Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><ct:Override PartName="/${folder}/sig1.xml" ContentType="${SIGNATURE_TYPE}xmlsignature+xml"/><ct:Override PartName="/${folder}/sig2.xml" ContentType="${SIGNATURE_TYPE}xmlsignature+xml"/></ct:Types>`,
	);
	zip.file(`${folder}/origin.sigs`, '');
	zip.file(
		`${folder}/_rels/origin.sigs.rels`,
		`<Relationships xmlns="${RELS}"><Relationship Id="s1" Type="${DIGITAL_SIGNATURE_REL_TYPE}" Target="sig1.xml"/><Relationship Id="s2" Type="${DIGITAL_SIGNATURE_REL_TYPE}" Target="sig2.xml"/></Relationships>`,
	);
	for (const name of ['sig1', 'sig2'])
		zip.file(`${folder}/${name}.xml`, '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>');
	// A certificate reached through a relationship, without its own content-type override.
	zip.file(
		`${folder}/_rels/sig1.xml.rels`,
		`<Relationships xmlns="${RELS}"><Relationship Id="cert" Type="${RELS}/digital-signature/certificate" Target="/certs/certificate%20one.cer"/></Relationships>`,
	);
	zip.file('certs/certificate one.cer', new Uint8Array([1, 2, 3]));
	zip.file('customXml/item1.xml', '<payload>keep this</payload>');
	zip.file('word/vbaProject.bin', new Uint8Array([4, 5, 6]));
	return zip.generateAsync({ type: 'uint8array' });
}

describe('DOCX package signatures', () => {
	it.each(['_xmlsignatures', 'custom/signing'])(
		'warns and preserves an untouched %s package',
		async (folder) => {
			const bytes = await signedDocument(folder);
			const loaded = await loadDocx(bytes);
			expect(loaded.model.warnings).toEqual(
				expect.arrayContaining([
					expect.stringMatching(/digital signatures.*not been verified.*Saving edits removes/),
				]),
			);
			expect(await loaded.save()).toEqual(bytes);
			expect(await loaded.save(structuredClone(loaded.model))).toEqual(bytes);
		},
	);

	it.each(['_xmlsignatures', 'custom/signing'])(
		'strips %s signatures after a text edit and preserves unrelated parts',
		async (folder) => {
			const original = await signedDocument(folder);
			const loaded = await loadDocx(original);
			const next = structuredClone(loaded.model);
			expectParagraph(next.blocks[0]).runs = [{ text: 'Edited document' }];
			const bytes = await loaded.save(next);
			const zip = await JSZip.loadAsync(bytes);
			expect(
				Object.keys(zip.files).filter((path) => !zip.files[path]!.dir && path.startsWith(folder)),
			).toEqual([]);
			expect(zip.file('certs/certificate one.cer')).toBeNull();
			expect(await must(zip.file('_rels/.rels')).async('string')).not.toContain(
				'digital-signature',
			);
			expect(await must(zip.file('[Content_Types].xml')).async('string')).not.toContain(
				'digital-signature',
			);
			expect(await must(zip.file('customXml/item1.xml')).async('string')).toBe(
				'<payload>keep this</payload>',
			);
			expect(await must(zip.file('word/vbaProject.bin')).async('uint8array')).toEqual(
				new Uint8Array([4, 5, 6]),
			);
			const reloaded = await loadDocx(bytes);
			expect(expectParagraph(reloaded.model.blocks[0]).runs[0]?.text).toBe('Edited document');
			expect(
				reloaded.model.warnings.some((warning) => warning.includes('digital signatures')),
			).toBe(false);
			// Saving does not mutate the source context or remove signatures from the original model.
			expect(await loaded.save()).toEqual(original);
		},
	);

	it('also strips signatures for edits outside the body', async () => {
		const loaded = await loadDocx(await signedDocument());
		loaded.model.properties = { title: 'Changed title' };
		const reloaded = await loadDocx(await loaded.save());
		expect(reloaded.model.properties?.title).toBe('Changed title');
		expect(reloaded.model.warnings.some((warning) => warning.includes('digital signatures'))).toBe(
			false,
		);
	});

	it('does not warn for an unsigned document', async () => {
		const loaded = await loadDocx(await saveDocx(createDocument()));
		expect(loaded.model.warnings.some((warning) => warning.includes('digital signatures'))).toBe(
			false,
		);
	});
});
