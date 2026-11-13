import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';

import { resolveReferenceUriToPart } from '../../../opc/signature/reference-utils';
import {
	type ParsedSignature,
	type SignatureDigestStatus,
	parseSignatureXml,
	verifySignatureDigests,
} from './signature-detection';
import { decodeXmlEntities } from './xml-entities';

type XmlObject = Record<string, unknown>;

/** A signature of a package with the integrity check of the parts it covers. */
export interface PackageSignatureInfo extends ParsedSignature {
	/**
	 * The referenced parts' digests against their current bytes. This is an integrity check only:
	 * the signature value, certificate chain, validity and revocation are not evaluated, so it says
	 * nothing about who signed.
	 */
	integrity: SignatureDigestStatus;
}

const SIGNATURE_PART = /^_xmlsignatures\/.+\.xml$/u;

function signatureParser(): XMLParser {
	// Mirrors the loader's parser: numeric character references in a certificate DN are decoded
	// (`CN=M&#xFC;ller CA`), and DTD handling is off on a path that reads attacker-supplied bytes.
	return new XMLParser({
		ignoreAttributes: false,
		attributeNamePrefix: '@_',
		parseAttributeValue: false,
		parseTagValue: false,
		processEntities: false,
		tagValueProcessor: (_tagName: string, tagValue: string) => decodeXmlEntities(tagValue),
		attributeValueProcessor: (_attrName: string, attrValue: string) => decodeXmlEntities(attrValue),
	});
}

async function readSignatures(zip: JSZip): Promise<ParsedSignature[]> {
	const parser = signatureParser();
	const result: ParsedSignature[] = [];
	for (const path of Object.keys(zip.files)) {
		const file = zip.file(path);
		if (file && SIGNATURE_PART.test(path)) {
			result.push(parseSignatureXml(parser.parse(await file.async('string')) as XmlObject, path));
		}
	}
	return result;
}

/**
 * Parse the digital signatures of a `.pptx` ZIP buffer (best effort: an empty list when there are
 * none or the package cannot be read).
 */
export async function parseSignaturesFromBuffer(
	buffer: ArrayBuffer | Uint8Array,
): Promise<ParsedSignature[]> {
	try {
		return await readSignatures(await JSZip.loadAsync(buffer));
	} catch {
		return [];
	}
}

/**
 * {@link parseSignaturesFromBuffer} plus a digest check of each signature's referenced parts.
 * Never throws; an unreadable package yields an empty list.
 */
export async function inspectSignaturesFromBuffer(
	buffer: ArrayBuffer | Uint8Array,
): Promise<PackageSignatureInfo[]> {
	try {
		const zip = await JSZip.loadAsync(buffer);
		const result: PackageSignatureInfo[] = [];
		for (const signature of await readSignatures(zip)) {
			const integrity = await verifySignatureDigests(signature, async (uri) => {
				const path = resolveReferenceUriToPart(uri.split('?')[0] ?? uri);
				const file = path ? zip.file(path) : null;
				return file ? await file.async('arraybuffer') : undefined;
			}).catch((): SignatureDigestStatus => 'unverified');
			result.push({ ...signature, integrity });
		}
		return result;
	} catch {
		return [];
	}
}
