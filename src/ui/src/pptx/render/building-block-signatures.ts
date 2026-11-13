import { inspectSignaturesFromBuffer } from 'ooxml-core/pptx';
import type { PackageSignatureInfo } from 'ooxml-core/pptx';

import {
	BUILDING_BLOCK_OK,
	buildingBlockFailure,
	type BuildingBlockResult,
} from './building-block-actions';

/**
 * Framework-free digital-signature workflow of a host-owned shell. It reports what the package
 * carries and whether the signed parts still match; it does not decide trust.
 *
 * Honest limits: the browser checks the digests of the parts a signature covers, not the signature
 * value, the certificate chain or revocation, so no signature is ever reported as trusted. The
 * viewer cannot create a signature (there is no signing key or certificate store here). Saving
 * always writes the package without its signatures, so "remove" means "the saved file will not
 * carry them".
 */

export interface HostSignatureInfo {
	/** The signature part in the package, for example `_xmlsignatures/sig1.xml`. */
	path: string;
	/** The signer certificate's subject, when the signature carries one. */
	signer?: string;
	issuer?: string;
	/** ISO dates from the certificate. */
	validFrom?: string;
	validTo?: string;
	/** The certificate's validity period has ended. */
	expired: boolean;
	/** The covered parts' digests against their current bytes (an integrity check only). */
	integrity: PackageSignatureInfo['integrity'];
}

export interface SignaturesState {
	/** The package carries at least one signature. */
	signed: boolean;
	count: number;
	/** Saving drops the signatures (always true: edits invalidate them and the viewer cannot re-sign). */
	removedOnSave: true;
	/** What the host can ask for. Creating a signature and a trust decision are not available. */
	supported: { list: true; remove: true; create: false; trust: false };
}

export function describeSignaturesState(input: {
	signed: boolean;
	count: number;
}): SignaturesState {
	return {
		signed: input.signed,
		count: input.count,
		removedOnSave: true,
		supported: { list: true, remove: true, create: false, trust: false },
	};
}

export function toHostSignatureInfo(signature: PackageSignatureInfo): HostSignatureInfo {
	const { certificate } = signature;
	return {
		path: signature.signaturePath,
		...(certificate?.subject ? { signer: certificate.subject } : {}),
		...(certificate?.issuer ? { issuer: certificate.issuer } : {}),
		...(certificate?.validFrom ? { validFrom: certificate.validFrom } : {}),
		...(certificate?.validTo ? { validTo: certificate.validTo } : {}),
		expired: signature.status === 'expired',
		integrity: signature.integrity,
	};
}

/** List the signatures of a package's bytes (empty for an unsigned or unreadable package). */
export async function listPackageSignatures(
	content: ArrayBuffer | Uint8Array | null | undefined,
): Promise<HostSignatureInfo[]> {
	if (!content) {
		return [];
	}
	return (await inspectSignaturesFromBuffer(content)).map(toHostSignatureInfo);
}

export interface SignatureActions {
	state: SignaturesState;
	/** Read the signatures from the loaded package. Allowed read-only and in Protected View. */
	list: () => Promise<HostSignatureInfo[]>;
	/**
	 * Drop the signatures from the file the next save writes (marks the deck changed). Fails
	 * with `no-signatures` when there is nothing to remove.
	 */
	remove: () => BuildingBlockResult;
	/** Not supported: the viewer cannot create a signature. Always refused. */
	sign: () => BuildingBlockResult;
}

export interface SignatureActionsInput {
	state: SignaturesState;
	content: () => ArrayBuffer | Uint8Array | null | undefined;
	refuse: () => BuildingBlockResult | undefined;
	markDirty: () => void;
	complete: (result: BuildingBlockResult) => BuildingBlockResult;
}

export function buildSignatureActions(input: SignatureActionsInput): SignatureActions {
	return {
		state: input.state,
		list: () => listPackageSignatures(input.content()),
		remove: () => {
			const denied = input.refuse();
			if (denied) {
				return denied;
			}
			if (!input.state.signed) {
				return buildingBlockFailure('no-signatures');
			}
			input.markDirty();
			return input.complete(BUILDING_BLOCK_OK);
		},
		sign: () => buildingBlockFailure('unsupported'),
	};
}
