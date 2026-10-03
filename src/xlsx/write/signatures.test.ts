import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from './index.js';

// A signed package built by hand the way Excel lays one out (Excel COM can only sign with a
// certificate from the user's store): a root origin relationship, the origin part, its
// relationships to each signature, and the content types for both.
const ORIGIN_REL =
	'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin';
const SIGNATURE_REL =
	'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature';

async function signedWorkbook(): Promise<Uint8Array> {
	const workbook = createWorkbook();
	putCell(workbook.sheets[0]!, 0, 0, { value: 'signed' });
	const zip = await JSZip.loadAsync(await saveXlsx(workbook));
	const rels = (await zip.file('_rels/.rels')!.async('string')).replace(
		'</Relationships>',
		`<Relationship Id="rIdSig" Type="${ORIGIN_REL}" Target="_xmlsignatures/origin.sigs"/></Relationships>`,
	);
	zip.file('_rels/.rels', rels);
	const types = (await zip.file('[Content_Types].xml')!.async('string')).replace(
		'<Default Extension="xml"',
		'<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/><Default Extension="xml"',
	);
	zip.file(
		'[Content_Types].xml',
		types.replace(
			'</Types>',
			'<Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/></Types>',
		),
	);
	zip.file('_xmlsignatures/origin.sigs', '');
	zip.file(
		'_xmlsignatures/_rels/origin.sigs.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${SIGNATURE_REL}" Target="sig1.xml"/></Relationships>`,
	);
	zip.file(
		'_xmlsignatures/sig1.xml',
		'<?xml version="1.0" encoding="UTF-8"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#" Id="idPackageSignature"><SignedInfo/><SignatureValue>AAAA</SignatureValue></Signature>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('digitally signed workbooks', () => {
	it('records the signatures and warns that saving removes them', async () => {
		const workbook = await loadXlsx(await signedWorkbook());
		expect(workbook.signatures).toStrictEqual({ count: 1, parts: ['_xmlsignatures/sig1.xml'] });
		expect(workbook.warnings.some((warning) => /digitally signed/.test(warning))).toBe(true);
		expect((await loadXlsx(await saveXlsx(createWorkbook()))).signatures).toBeUndefined();
	});

	it('strips the signature parts, relationships and content types on save', async () => {
		const workbook = await loadXlsx(await signedWorkbook());
		const zip = await JSZip.loadAsync(await saveXlsx(workbook));
		expect(Object.keys(zip.files).filter((name) => name.startsWith('_xmlsignatures'))).toEqual([]);
		const rels = await zip.file('_rels/.rels')!.async('string');
		expect(rels).not.toContain('digital-signature');
		const types = await zip.file('[Content_Types].xml')!.async('string');
		expect(types).not.toContain('digital-signature');
		expect(types).not.toContain('sigs');
		const reloaded = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(reloaded.signatures).toBeUndefined();
		expect(getCell(reloaded.sheets[0]!, 0, 0)?.value).toBe('signed');
	});
});
