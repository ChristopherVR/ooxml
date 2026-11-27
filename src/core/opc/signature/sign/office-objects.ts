// The `Object`s an Office package signature carries besides the part manifest (ECMA-376 Part 2,
// clause 13; [MS-OFFCRYPTO] 2.5.2, "XAdES" and "SignatureInfoV1"): the Office details that tie a
// signature to a signature line, and the XAdES signed properties naming the signing certificate.
import { escapeXmlText } from '../xml-utils';
import { OFFICE_DIGSIG_NS, OFFICE_OBJECT_ID } from './office-ids';

export { OFFICE_DIGSIG_NS, OFFICE_OBJECT_ID };

export const XADES_NS = 'http://uri.etsi.org/01903/v1.3.2#';
export const SIGNATURE_ID = 'idPackageSignature';
export const PACKAGE_OBJECT_ID = 'idPackageObject';
export const SIGNED_PROPERTIES_ID = 'idSignedProperties';
const NO_PROVIDER = '{00000000-0000-0000-0000-000000000000}';

/** What a signer adds to the signature beyond the package content. */
export interface OfficeSignatureDetails {
	/** The `{GUID}` of the signature line this signature signs (`o:signatureline/@id`). */
	signatureLineId?: string;
	/** The name the signer typed, shown on the signed line. */
	signatureText?: string;
	/** A picture of the signer's signature (Base64 PNG), shown on the signed line instead of text. */
	signatureImage?: string;
	/** "Purpose for signing this document", or the signer's comments on a signature line. */
	comments?: string;
	/** XAdES `ClaimedRole`: the signer's title. */
	signerRole?: string;
}

/** `Object Id="idOfficeObject"`: `SignatureInfoV1`, whose element order the schema fixes. */
export function officeObjectXml(details: OfficeSignatureDetails): string {
	const text = (name: string, value = '') => `<${name}>${escapeXmlText(value)}</${name}>`;
	const info = [
		text('SetupID', details.signatureLineId),
		text('SignatureText', details.signatureImage ? '' : details.signatureText),
		text('SignatureImage', details.signatureImage),
		text('SignatureComments', details.comments),
		// The signing environment Office records. They describe the format version written, not
		// the machine; Office reads them only to describe the signature.
		text('WindowsVersion', '10.0'),
		text('OfficeVersion', '16.0'),
		text('ApplicationVersion', '16.0'),
		text('Monitors', '1'),
		text('HorizontalResolution', '1920'),
		text('VerticalResolution', '1080'),
		text('ColorDepth', '32'),
		text('SignatureProviderId', NO_PROVIDER),
		text('SignatureProviderUrl'),
		text('SignatureProviderDetails', '9'),
		// 2: the signature signs a signature line; 1: an invisible signature.
		text('SignatureType', details.signatureLineId ? '2' : '1'),
	].join('');
	return (
		`<Object Id="${OFFICE_OBJECT_ID}"><SignatureProperties>` +
		`<SignatureProperty Id="idOfficeV1Details" Target="#${SIGNATURE_ID}">` +
		`<SignatureInfoV1 xmlns="${OFFICE_DIGSIG_NS}">${info}</SignatureInfoV1>` +
		'</SignatureProperty></SignatureProperties></Object>'
	);
}

/** The certificate as XAdES names it: its digest, issuer and decimal serial number. */
export interface SigningCertificateRef {
	digestMethod: string;
	digestValue: string;
	issuer: string;
	serialNumber: string;
}

/** The XAdES `QualifyingProperties` object (XAdES-BES), signed through `#idSignedProperties`. */
export function qualifyingPropertiesXml(
	signingTime: string,
	certificate: SigningCertificateRef,
	details: OfficeSignatureDetails,
): string {
	const role = details.signerRole
		? `<xd:SignerRole><xd:ClaimedRoles><xd:ClaimedRole>${escapeXmlText(details.signerRole)}</xd:ClaimedRole></xd:ClaimedRoles></xd:SignerRole>`
		: '';
	const commitment = details.comments
		? '<xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId>' +
			'<xd:Identifier>http://uri.etsi.org/01903/v1.2.2#ProofOfApproval</xd:Identifier>' +
			'<xd:Description>Approved this document</xd:Description></xd:CommitmentTypeId>' +
			'<xd:AllSignedDataObjects/><xd:CommitmentTypeQualifiers>' +
			`<xd:CommitmentTypeQualifier>${escapeXmlText(details.comments)}</xd:CommitmentTypeQualifier>` +
			'</xd:CommitmentTypeQualifiers></xd:CommitmentTypeIndication></xd:SignedDataObjectProperties>'
		: '';
	return (
		`<Object><xd:QualifyingProperties xmlns:xd="${XADES_NS}" Target="#${SIGNATURE_ID}">` +
		`<xd:SignedProperties Id="${SIGNED_PROPERTIES_ID}"><xd:SignedSignatureProperties>` +
		`<xd:SigningTime>${signingTime}</xd:SigningTime>` +
		'<xd:SigningCertificate><xd:Cert><xd:CertDigest>' +
		`<DigestMethod Algorithm="${certificate.digestMethod}"/>` +
		`<DigestValue>${certificate.digestValue}</DigestValue></xd:CertDigest>` +
		`<xd:IssuerSerial><X509IssuerName>${escapeXmlText(certificate.issuer)}</X509IssuerName>` +
		`<X509SerialNumber>${certificate.serialNumber}</X509SerialNumber></xd:IssuerSerial>` +
		'</xd:Cert></xd:SigningCertificate>' +
		'<xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier>' +
		`${role}</xd:SignedSignatureProperties>${commitment}</xd:SignedProperties>` +
		'</xd:QualifyingProperties></Object>'
	);
}
