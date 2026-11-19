import JSZip from 'jszip';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

import {
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
	const sign = async () => {
		const material = new TextEncoder().encode(signer.keyPem + signer.certPem);
		const bytes = await createTestPackage().generateAsync({ type: 'uint8array' });
		return signPptxWithCertificate(bytes, material, { certificatePath: 'signer.pem' });
	};

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
});
