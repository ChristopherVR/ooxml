// Reads which signature line a verified signature signs (`SignatureInfoV1/SetupID`) and what the
// signer put on it, so a product can show the line as signed. Only a signature the verifier
// accepted is read, and only from its `idOfficeObject`, which that signature's digest protects.
import type JSZip from 'jszip';
import { decodeXmlBytes } from '../../../xml/index';
import { XMLDSIG_NS } from '../constants';
import { childElements, localNameOf, parseStrictXml } from '../verify/dom';
import { SIGNATURE_LIMITS } from '../verify/limits';
import { isVerifiedSignatureState, type SignatureSummary } from '../state';
import { OFFICE_DIGSIG_NS, OFFICE_OBJECT_ID } from './office-objects';

const MAX_TEXT = 255;
const MAX_IMAGE = 2 * 1024 * 1024;

function officeInfo(xml: string): { setupId: string; text?: string; image?: string } | undefined {
	const root = parseStrictXml(xml).documentElement;
	if (root.namespaceURI !== XMLDSIG_NS || localNameOf(root) !== 'Signature') return undefined;
	const signedInfo = childElements(root).find((child) => localNameOf(child) === 'SignedInfo');
	const protectedObject = signedInfo
		? childElements(signedInfo).some(
				(reference) => reference.getAttribute('URI') === `#${OFFICE_OBJECT_ID}`,
			)
		: false;
	const object = childElements(root).find(
		(child) => localNameOf(child) === 'Object' && child.getAttribute('Id') === OFFICE_OBJECT_ID,
	);
	if (!protectedObject || !object) return undefined;
	const info = object.getElementsByTagNameNS(OFFICE_DIGSIG_NS, 'SignatureInfoV1')[0];
	if (!info) return undefined;
	const value = (name: string) =>
		info.getElementsByTagNameNS(OFFICE_DIGSIG_NS, name)[0]?.textContent?.trim() ?? '';
	const setupId = value('SetupID');
	if (!/^\{[0-9A-Fa-f-]{36}\}$/u.test(setupId)) return undefined;
	const text = value('SignatureText').slice(0, MAX_TEXT);
	const image = value('SignatureImage').replace(/\s+/gu, '');
	return {
		setupId,
		...(text ? { text } : {}),
		...(image && image.length <= MAX_IMAGE && /^[A-Za-z0-9+/]+={0,2}$/u.test(image)
			? { image }
			: {}),
	};
}

/**
 * The summary with each verified signature that signs a signature line given the line's id and
 * what the signer wrote or drew on it. Unreadable signature parts are left as they are.
 */
export async function withSignatureLineSignings(
	zip: JSZip,
	summary: SignatureSummary,
): Promise<SignatureSummary> {
	const signatures = await Promise.all(
		summary.signatures.map(async (entry) => {
			if (!isVerifiedSignatureState(entry.state)) return entry;
			const file = zip.file(entry.path);
			if (!file) return entry;
			try {
				const bytes = await file.async('uint8array');
				if (bytes.byteLength > SIGNATURE_LIMITS.maxSignaturePartBytes) return entry;
				const info = officeInfo(decodeXmlBytes(bytes));
				return info
					? {
							...entry,
							signatureLineId: info.setupId,
							...(info.text ? { signatureText: info.text } : {}),
							...(info.image ? { signatureImage: info.image } : {}),
						}
					: entry;
			} catch {
				return entry;
			}
		}),
	);
	return { ...summary, signatures };
}
