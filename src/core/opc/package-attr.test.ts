import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { RELATIONSHIP_TYPES } from './relationship-types';
import { ensureContentTypeOverride, ensureRelationship, findOfficeDocumentPart } from './package';

const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

describe('OPC attribute matching', () => {
	it('does not read Id from inside an attribute such as rId', async () => {
		const zip = new JSZip();
		zip.file(
			'word/_rels/document.xml.rels',
			`<Relationships xmlns="${NS_RELS}"><Relationship rId="stale" Id="rId7" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/></Relationships>`,
		);
		expect(
			await ensureRelationship(
				zip,
				'word/_rels/document.xml.rels',
				RELATIONSHIP_TYPES.styles,
				'styles.xml',
			),
		).toBe('rId7');
	});

	it('numbers a new relationship after the real Id, not a value found in rId', async () => {
		const zip = new JSZip();
		zip.file(
			'word/_rels/document.xml.rels',
			`<Relationships xmlns="${NS_RELS}"><Relationship rId="rId99" Id="rId2" Type="urn:other" Target="a.xml"/></Relationships>`,
		);
		expect(
			await ensureRelationship(
				zip,
				'word/_rels/document.xml.rels',
				RELATIONSHIP_TYPES.theme,
				't.xml',
			),
		).toBe('rId3');
	});

	it('accepts single-quoted attributes when reusing a relationship', async () => {
		const zip = new JSZip();
		zip.file(
			'word/_rels/document.xml.rels',
			`<Relationships xmlns="${NS_RELS}"><Relationship Id='rId4' Type='${RELATIONSHIP_TYPES.numbering}' Target='numbering.xml'/></Relationships>`,
		);
		expect(
			await ensureRelationship(
				zip,
				'word/_rels/document.xml.rels',
				RELATIONSHIP_TYPES.numbering,
				'numbering.xml',
			),
		).toBe('rId4');
	});

	it('does not duplicate a single-quoted content type override', async () => {
		const zip = new JSZip();
		zip.file(
			'[Content_Types].xml',
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName='/word/numbering.xml' ContentType='numbering+xml'/></Types>`,
		);
		await ensureContentTypeOverride(zip, 'word/numbering.xml', 'numbering+xml');
		const xml = await zip.file('[Content_Types].xml')!.async('string');
		expect(xml.match(/numbering\.xml/g)).toHaveLength(1);
	});

	it('reads the office document target from single-quoted package relationships', async () => {
		const zip = new JSZip();
		zip.file(
			'_rels/.rels',
			`<Relationships xmlns="${NS_RELS}"><Relationship Id='rId1' Type='${RELATIONSHIP_TYPES.officeDocument}' Target='word/document.xml'/></Relationships>`,
		);
		expect(await findOfficeDocumentPart(zip)).toBe('word/document.xml');
	});
});
