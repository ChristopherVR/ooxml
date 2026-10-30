import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { DIGITAL_SIGNATURE_ORIGIN_REL_TYPE, DIGITAL_SIGNATURE_REL_TYPE } from './constants';
import { inspectPackageSignatures, stripPackageSignatures } from './package';

const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

describe('package signature metadata', () => {
	it('leaves unsigned metadata and parts unchanged', async () => {
		const zip = new JSZip();
		const rels = `<Relationships xmlns="${RELS}" />`;
		const types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" />';
		zip.file('_rels/.rels', rels);
		zip.file('[Content_Types].xml', types);
		expect((await inspectPackageSignatures(zip)).hasSignatures).toBe(false);
		await stripPackageSignatures(zip);
		expect(await zip.file('_rels/.rels')!.async('string')).toBe(rels);
		expect(await zip.file('[Content_Types].xml')!.async('string')).toBe(types);
	});

	it('removes incomplete and externally targeted signature metadata without fetching it', async () => {
		const zip = new JSZip();
		zip.file(
			'_rels/.rels',
			`<Relationships xmlns="${RELS}"><Relationship Id="sig" Type="${DIGITAL_SIGNATURE_ORIGIN_REL_TYPE}" Target="https://example.invalid/origin.sigs" TargetMode="External"/></Relationships>`,
		);
		expect((await inspectPackageSignatures(zip)).hasSignatures).toBe(true);
		await stripPackageSignatures(zip);
		expect((await inspectPackageSignatures(zip)).hasSignatures).toBe(false);
	});

	it('follows custom signature relationships once even when cyclic', async () => {
		const zip = new JSZip();
		zip.file(
			'_rels/.rels',
			`<Relationships xmlns="${RELS}"><Relationship Id="sig" Type="${DIGITAL_SIGNATURE_ORIGIN_REL_TYPE}" Target="signing/origin"/></Relationships>`,
		);
		zip.file('signing/origin', '');
		zip.file(
			'signing/_rels/origin.rels',
			`<Relationships xmlns="${RELS}"><Relationship Id="sig" Type="${DIGITAL_SIGNATURE_REL_TYPE}" Target="origin"/></Relationships>`,
		);
		await stripPackageSignatures(zip);
		expect(zip.file('signing/origin')).toBeNull();
		expect(zip.file('signing/_rels/origin.rels')).toBeNull();
		expect((await inspectPackageSignatures(zip)).hasSignatures).toBe(false);
	});

	it('removes orphan signature parts with mixed-case paths', async () => {
		const zip = new JSZip();
		zip.file('_XMLSIGNATURES/SIG1.XML', '<Signature/>');
		expect((await inspectPackageSignatures(zip)).hasSignatures).toBe(true);
		await stripPackageSignatures(zip);
		expect((await inspectPackageSignatures(zip)).hasSignatures).toBe(false);
	});
});
