import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

import { PptxHandler } from '../core/PptxHandler';
import {
	ENTERPRISE_REQUIRE_REVOCATION_ENV,
	ENTERPRISE_REQUIRE_TIMESTAMP_ENV,
	ENTERPRISE_TRUST_ROOTS_PEM_ENV,
} from '../core/utils/signature-constants';
import type { DigitalSignatureReport, SignatureDetail } from '../core/utils/signature-types';
import {
	createManifestReferences,
	createOfficeSignatureXml,
	createTestPackage,
	createTestParty,
	CUSTOM_XML_RELATIONSHIP,
	partReference,
	relationshipXml,
	SLIDE_CONTENT_TYPE,
	SLIDE_PART,
	SLIDE_RELATIONSHIP,
	withSignature,
	type TestParty,
} from '../test-support/signature-package';
import { inspectPptxDigitalSignatures } from './inspection';
import { signPptxWithCertificate } from './signing';

const PRESENTATION_RELS = 'ppt/_rels/presentation.xml.rels';
const CONTENT_TYPES = '[Content_Types].xml';
const SIGNATURE_PART = '_xmlsignatures/sig1.xml';
const RELATIONSHIPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const HYPERLINK_RELATIONSHIP =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const BOM = '\uFEFF';
const BOM_DECK = fileURLToPath(
	new URL('../__tests__/fixtures/e2e/absolute-path-rels.pptx', import.meta.url),
);
const CORE_PROPERTIES =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

let root: TestParty;
let signer: TestParty;

beforeAll(() => {
	root = createTestParty({ name: 'Test Root', ca: true });
	signer = createTestParty({
		name: 'Signer',
		issuer: root,
		extensions: [{ name: 'extKeyUsage', emailProtection: true }],
	});
});

beforeEach(() => {
	vi.stubEnv(ENTERPRISE_TRUST_ROOTS_PEM_ENV, root.certPem);
});

afterEach(() => {
	vi.unstubAllEnvs();
});

type Mutation = (zip: JSZip) => void | Promise<void>;

interface PackageOptions {
	selectBy?: 'id' | 'type';
	/** Runs before signing, so the change is part of what is signed. */
	prepare?: Mutation;
	references?: (zip: JSZip) => Promise<string[]>;
	unsignedSuffix?: string;
}

/** Sign the test package, apply `tamper` afterwards and inspect the result. */
async function inspectSigned(
	tamper: Mutation = () => undefined,
	{ selectBy, prepare, references, unsignedSuffix }: PackageOptions = {},
): Promise<DigitalSignatureReport> {
	const zip = createTestPackage();
	await prepare?.(zip);
	const manifestReferences = references
		? await references(zip)
		: await createManifestReferences(zip, selectBy ? { selectBy } : {});
	await withSignature(
		zip,
		createOfficeSignatureXml({
			signer,
			manifestReferences,
			...(unsignedSuffix === undefined ? {} : { unsignedSuffix }),
		}),
	);
	await tamper(zip);
	return inspectPptxDigitalSignatures(await zip.generateAsync({ type: 'uint8array' }));
}

const detailOf = (report: DigitalSignatureReport): SignatureDetail => report.details![0]!;

const editPart = (part: string, edit: (xml: string) => string): Mutation => {
	return async (zip) => {
		zip.file(part, edit(await zip.file(part)!.async('string')));
	};
};

const appendRelationship = (part: string, xml: string): Mutation =>
	editPart(part, (rels) => rels.replace('</Relationships>', `${xml}</Relationships>`));

const VERIFIED = ['verified-trusted', 'verified-untrusted'];

describe('relationships the verifier and the loader could read differently', () => {
	const slideElsewhere = `Type="${SLIDE_RELATIONSHIP}" Target="slides/other.xml"`;

	it.each<[string, Mutation, PackageOptions?]>([
		[
			'a Relationship element in another namespace',
			appendRelationship(
				PRESENTATION_RELS,
				`<Relationship xmlns="urn:other" Id="rId1" ${slideElsewhere}/>`,
			),
		],
		[
			'a prefixed Relationship element in another namespace',
			appendRelationship(
				PRESENTATION_RELS,
				`<o:Relationship xmlns:o="urn:other" Id="rId1" ${slideElsewhere}/>`,
			),
		],
		[
			'a repeated Id with the signed type',
			appendRelationship(PRESENTATION_RELS, `<Relationship Id="rId1" ${slideElsewhere}/>`),
		],
		[
			'a repeated Id with another type, signed by Id',
			appendRelationship(
				PRESENTATION_RELS,
				relationshipXml('rId1', 'urn:other', 'slides/other.xml'),
			),
		],
		[
			'a repeated Id with another type, signed by type',
			appendRelationship(
				PRESENTATION_RELS,
				relationshipXml('rId1', 'urn:other', 'slides/other.xml'),
			),
			{ selectBy: 'type' },
		],
		[
			'a target mode in another case on a signed relationship',
			editPart(PRESENTATION_RELS, (rels) =>
				rels.replace('Target="slides', 'TargetMode="external" Target="slides'),
			),
		],
	])('does not verify %s', async (_label, tamper, options) => {
		const report = await inspectSigned(tamper, options);
		expect(VERIFIED).not.toContain(report.verificationStatus);
	});
});

describe('stored items no signature covers', () => {
	it.each<[string, Mutation, string]>([
		[
			'a case variant of a signed part',
			(zip) => void zip.file('ppt/slides/SLIDE1.xml', '<sld/>'),
			'ppt/slides/SLIDE1.xml',
		],
		[
			'a percent-encoded variant of a signed part',
			(zip) => void zip.file('ppt/slides/slide%31.xml', '<sld/>'),
			'ppt/slides/slide%31.xml',
		],
		[
			'a case variant of a relationships part',
			(zip) => void zip.file('ppt/slides/_RELS/slide1.xml.rels', '<Relationships/>'),
			'ppt/slides/_RELS/slide1.xml.rels',
		],
		[
			'a part nothing refers to',
			(zip) => void zip.file('ppt/theme/theme9.xml', '<theme/>'),
			'ppt/theme/theme9.xml',
		],
	])('reports %s', async (_label, tamper, part) => {
		const report = await inspectSigned(tamper);
		expect(report.verificationStatus).toBe('coverage-incomplete');
		expect(detailOf(report).coverageGaps.map((gap) => gap.part)).toContain(part);
	});

	it.each<[string, string]>([
		[
			'a type that only ends like an exempt one',
			relationshipXml('rId9', 'urn:x/customXml', 'x.xml'),
		],
		[
			'an exempt type aimed at document content',
			relationshipXml('rId9', CUSTOM_XML_RELATIONSHIP, 'slides/other.xml'),
		],
		[
			'an exempt type in another case',
			relationshipXml('rId9', CUSTOM_XML_RELATIONSHIP.toLowerCase(), '../customXml/item1.xml'),
		],
		[
			'an exempt type with an external target',
			`<Relationship Id="rId9" Type="${CORE_PROPERTIES}" Target="../docProps/x.xml" TargetMode="External"/>`,
		],
	])('reports an unsigned relationship with %s', async (_label, xml) => {
		const report = await inspectSigned(appendRelationship(PRESENTATION_RELS, xml));
		expect(report.verificationStatus).toBe('coverage-incomplete');
		expect(detailOf(report).coverageGaps).toContainEqual({
			kind: 'unsigned-relationship',
			part: PRESENTATION_RELS,
			relationshipId: 'rId9',
		});
	});

	it('still lets custom XML stay unsigned where Office leaves it', async () => {
		const report = await inspectSigned(async (zip) => {
			await appendRelationship(
				PRESENTATION_RELS,
				relationshipXml('rId9', CUSTOM_XML_RELATIONSHIP, '../customXml/item1.xml'),
			)(zip);
			zip.file('customXml/item1.xml', '<data/>');
			zip.file('customXml/itemProps1.xml', '<props/>');
			zip.file('customXml/_rels/item1.xml.rels', '<Relationships/>');
		});
		expect(report.verificationStatus).toBe('verified-trusted');
	});

	it('ignores items OPC does not map to parts', async () => {
		const report = await inspectSigned((zip) => void zip.file('[trash]/0000.dat', 'x'));
		expect(report.verificationStatus).toBe('verified-trusted');
	});
});

/** Signs the presentation relationships whole, so every relationship in them is selected. */
async function signingWholePresentationRelationships(zip: JSZip): Promise<string[]> {
	const references = await createManifestReferences(zip);
	references[1] = await partReference(zip, PRESENTATION_RELS, RELATIONSHIPS_CONTENT_TYPE);
	return references;
}

describe('signed relationships to parts a signature may leave out', () => {
	const addTarget = (type: string, target: string, part: string): Mutation => {
		return async (zip) => {
			await appendRelationship(PRESENTATION_RELS, relationshipXml('rId9', type, target))(zip);
			zip.file(part, '<data/>');
		};
	};
	const inspectWith = (prepare: Mutation) =>
		inspectSigned(() => undefined, {
			prepare,
			references: signingWholePresentationRelationships,
		});

	it('accepts custom XML that stays unsigned, and still lists it', async () => {
		const report = await inspectWith(
			addTarget(CUSTOM_XML_RELATIONSHIP, '../customXml/item1.xml', 'customXml/item1.xml'),
		);
		expect(report.verificationStatus).toBe('verified-trusted');
		expect(detailOf(report).coverageGaps).toEqual([]);
		expect(detailOf(report).uncoveredParts).toContain('customXml/item1.xml');
	});

	it.each([
		[
			'an exempt type aimed at document content',
			CUSTOM_XML_RELATIONSHIP,
			'slides/x.xml',
			'ppt/slides/x.xml',
		],
		[
			'a content type aimed at an exempt location',
			SLIDE_RELATIONSHIP,
			'../customXml/x.xml',
			'customXml/x.xml',
		],
	])('reports %s', async (_label, type, target, part) => {
		const report = await inspectWith(addTarget(type, target, part));
		expect(report.verificationStatus).toBe('coverage-incomplete');
		expect(detailOf(report).coverageGaps).toContainEqual({
			kind: 'unsigned-part',
			part,
			relationshipId: 'rId9',
		});
	});
});

describe('byte order marks', () => {
	const addMark = (part: string): Mutation => editPart(part, (xml) => BOM + xml);

	it.each([
		['the package relationships', '_rels/.rels'],
		['the presentation relationships', PRESENTATION_RELS],
		['the content types stream', CONTENT_TYPES],
	])('verifies a signed package with one on %s', async (_label, part) => {
		const report = await inspectSigned(() => undefined, { prepare: addMark(part) });
		expect(report.verificationStatus).toBe('verified-trusted');
	});

	it('verifies a signature part that starts with one', async () => {
		const report = await inspectSigned(addMark(SIGNATURE_PART));
		expect(report.verificationStatus).toBe('verified-trusted');
	});

	it('still fails a signed part changed behind one', async () => {
		const report = await inspectSigned(
			editPart(SLIDE_PART, (xml) => BOM + xml.replace('original', 'changed')),
		);
		expect(report.verificationStatus).toBe('digest-mismatch');
	});
});

describe('relationship values the schema allows', () => {
	const emptyTarget = `<Relationship Id="rId7" Type="${HYPERLINK_RELATIONSHIP}" Target="" TargetMode="External"/>`;
	const caseVariant = relationshipXml('RID1', HYPERLINK_RELATIONSHIP, 'slides/slide1.xml');

	it.each([
		['an empty external Target', emptyTarget],
		['an empty internal Target', relationshipXml('rId7', HYPERLINK_RELATIONSHIP, '')],
		['Ids that differ only in case', caseVariant],
	])('verifies %s', async (_label, xml) => {
		const report = await inspectSigned(() => undefined, {
			prepare: appendRelationship(PRESENTATION_RELS, xml),
			references: signingWholePresentationRelationships,
		});
		expect(report.verificationStatus).toBe('verified-trusted');
	});

	it('does not verify a case variant of a signed Id added afterwards', async () => {
		const report = await inspectSigned(appendRelationship(PRESENTATION_RELS, caseVariant));
		expect(VERIFIED).not.toContain(report.verificationStatus);
	});

	it('reports an unsigned relationship whose Id is a case variant of a signed one', async () => {
		const report = await inspectSigned(
			appendRelationship(PRESENTATION_RELS, relationshipXml('RID1', 'urn:other', 'slides/x.xml')),
			{ selectBy: 'type' },
		);
		expect(report.verificationStatus).toBe('coverage-incomplete');
		expect(detailOf(report).coverageGaps).toContainEqual({
			kind: 'unsigned-relationship',
			part: PRESENTATION_RELS,
			relationshipId: 'RID1',
		});
	});
});

describe('the revocation policy', () => {
	beforeEach(() => {
		vi.stubEnv(ENTERPRISE_REQUIRE_REVOCATION_ENV, 'true');
	});

	it('marks an intact signature whose revocation is unknown as untrusted', async () => {
		const report = await inspectSigned();
		expect(report.verificationStatus).toBe('verified-untrusted');
		expect(detailOf(report).status).toBe('certificate-untrusted');
	});

	it('does not hide a coverage gap', async () => {
		const report = await inspectSigned((zip) => void zip.file('ppt/slides/x.xml', '<sld/>'));
		expect(report.verificationStatus).toBe('coverage-incomplete');
	});

	it('does not hide a changed part', async () => {
		const report = await inspectSigned((zip) => void zip.file(SLIDE_PART, '<sld>changed</sld>'));
		expect(report.verificationStatus).toBe('digest-mismatch');
	});
});

describe('valid packages', () => {
	it('verifies when an Override names the part in another case', async () => {
		const report = await inspectSigned(() => undefined, {
			prepare: editPart(CONTENT_TYPES, (types) =>
				types.replace(`/${SLIDE_PART}`, '/PPT/Slides/Slide1.xml'),
			),
		});
		expect(report.verificationStatus).toBe('verified-trusted');
	});

	it('verifies a part whose stored name contains percent-encoding', async () => {
		const part = 'ppt/media/a%20b.xml';
		const report = await inspectSigned(() => undefined, {
			prepare: (zip) => void zip.file(part, '<x/>'),
			references: async (zip) => [
				...(await createManifestReferences(zip)),
				await partReference(zip, part, 'application/xml'),
			],
		});
		expect(report.verificationStatus).toBe('verified-trusted');
		expect(detailOf(report).uncoveredParts).toEqual(['docProps/core.xml']);
	});

	it('verifies a reference that spells the part name in another case', async () => {
		const report = await inspectSigned(() => undefined, {
			references: async (zip) => {
				const references = await createManifestReferences(zip);
				const slide = references.findIndex((reference) => reference.includes(SLIDE_PART));
				references[slide] = await partReference(
					zip,
					SLIDE_PART,
					SLIDE_CONTENT_TYPE,
					`/${SLIDE_PART.toUpperCase()}`,
				);
				return references;
			},
		});
		expect(report.verificationStatus).toBe('verified-trusted');
	});
});

describe('timestamps', () => {
	const unsignedToken = '<Object><EncapsulatedTimeStamp>MAMCAQE=</EncapsulatedTimeStamp></Object>';

	it('does not let an unvalidated token satisfy a policy that requires a timestamp', async () => {
		vi.stubEnv(ENTERPRISE_REQUIRE_TIMESTAMP_ENV, 'true');
		const report = await inspectSigned(() => undefined, { unsignedSuffix: unsignedToken });
		expect(report.verificationStatus).toBe('timestamp-untrusted');
		expect(detailOf(report).timestampAuthorityStatus).toBe('not-checked');
	});

	it('leaves a signature verified when no policy requires a timestamp', async () => {
		const report = await inspectSigned(() => undefined, { unsignedSuffix: unsignedToken });
		expect(report.verificationStatus).toBe('verified-trusted');
	});
});

describe('signatures this library writes', () => {
	const signBytes = (bytes: Uint8Array) =>
		signPptxWithCertificate(bytes, new TextEncoder().encode(signer.keyPem + signer.certPem), {
			certificatePath: 'signer.pem',
		});
	const signPackage = async (prepare: Mutation = () => undefined) => {
		const zip = createTestPackage();
		await prepare(zip);
		return signBytes(await zip.generateAsync({ type: 'uint8array' }));
	};
	const sign = () => signPackage();
	const relationshipsWith = (xml: string) => appendRelationship(PRESENTATION_RELS, xml);

	it('verify after signing', async () => {
		expect((await sign()).report.verificationStatus).toBe('verified-trusted');
	});

	it('bind each part to its declared content type', async () => {
		const { signedData } = await sign();
		const zip = await JSZip.loadAsync(signedData!);
		await editPart(CONTENT_TYPES, (types) =>
			types.replace(SLIDE_CONTENT_TYPE, 'application/octet-stream'),
		)(zip);
		const report = await inspectPptxDigitalSignatures(
			await zip.generateAsync({ type: 'uint8array' }),
		);
		expect(report.verificationStatus).toBe('digest-mismatch');
	});

	it('signs a deck whose parts start with byte order marks, and it still opens', async () => {
		const bytes = new Uint8Array(readFileSync(BOM_DECK));
		const result = await signBytes(bytes);
		expect(result.error).toBeUndefined();
		expect(result.report.verificationStatus).toBe('verified-trusted');
		const original = await new PptxHandler().load(bytes.slice().buffer);
		const signed = await new PptxHandler().load(result.signedData!.slice().buffer);
		expect(signed.slides).toHaveLength(original.slides.length);
	});

	it.each<[string, Mutation]>([
		[
			'an empty Target',
			relationshipsWith(
				`<Relationship Id="rId7" Type="${HYPERLINK_RELATIONSHIP}" Target="" TargetMode="External"/>`,
			),
		],
		[
			'Ids that differ only in case',
			relationshipsWith(relationshipXml('RID1', HYPERLINK_RELATIONSHIP, 'slides/slide1.xml')),
		],
		['an item OPC does not map to a part', (zip) => void zip.file('[trash]/0000.dat', 'x')],
		[
			'a leftover signature in another case',
			(zip) => void zip.file('_XMLSIGNATURES/old.xml', '<x/>'),
		],
	])('signs a package with %s', async (_label, prepare) => {
		const result = await signPackage(prepare);
		expect(result.error).toBeUndefined();
		expect(result.report.verificationStatus).toBe('verified-trusted');
		expect(result.report.signatureCount).toBe(1);
	});

	it('keeps one origin relationship when a package is signed again', async () => {
		const first = await sign();
		const second = await signBytes(first.signedData!);
		expect(second.report.verificationStatus).toBe('verified-trusted');
		const zip = await JSZip.loadAsync(second.signedData!);
		const rels = await zip.file('_rels/.rels')!.async('string');
		expect(rels.match(/digital-signature\/origin/g)).toHaveLength(1);
	});

	it('still succeeds when a policy withholds trust from the new signature', async () => {
		vi.stubEnv(ENTERPRISE_REQUIRE_TIMESTAMP_ENV, 'true');
		vi.stubEnv(ENTERPRISE_REQUIRE_REVOCATION_ENV, 'true');
		const result = await sign();
		expect(result.success).toBe(true);
		expect(result.report.verificationStatus).toBe('verified-untrusted');
	});

	it.each<[string, Mutation, string]>([
		[
			'a repeated relationship Id',
			relationshipsWith(relationshipXml('rId1', SLIDE_RELATIONSHIP, 'slides/slide1.xml')),
			PRESENTATION_RELS,
		],
		[
			'a relationship in another namespace',
			relationshipsWith(`<Relationship xmlns="urn:other" Id="rId2" Type="t" Target="a"/>`),
			PRESENTATION_RELS,
		],
		[
			'malformed relationships nothing refers to',
			(zip) => void zip.file('ppt/other/_rels/a.xml.rels', '<Relationships'),
			'ppt/other/_rels/a.xml.rels',
		],
		[
			'a space in a part name',
			(zip) => void zip.file('ppt/media/a b.png', 'x'),
			'ppt/media/a b.png',
		],
		[
			'a hash in a part name',
			(zip) => void zip.file('ppt/media/a#1.png', 'x'),
			'ppt/media/a#1.png',
		],
		[
			'a question mark in a part name',
			(zip) => void zip.file('ppt/media/a?.png', 'x'),
			'ppt/media/a?.png',
		],
		[
			'brackets in a part name',
			(zip) => void zip.file('ppt/media/a[1].png', 'x'),
			'ppt/media/a[1].png',
		],
		['a part name ending in a dot', (zip) => void zip.file('ppt/media/a.', 'x'), 'ppt/media/a.'],
	])('refuses to sign a package with %s', async (_label, prepare, part) => {
		const result = await signPackage(prepare);
		expect(result.success).toBe(false);
		expect(result.signedData).toBeUndefined();
		expect(result.report.verificationStatus).toBe('error');
		expect(result.error).toContain(part);
	});

	it('refuses a package whose signature could not reach any content', async () => {
		const result = await signPackage((zip) => void zip.remove('_rels/.rels'));
		expect(result.success).toBe(false);
		expect(result.signedData).toBeUndefined();
	});
});
