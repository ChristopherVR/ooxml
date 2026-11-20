import { inspectSignaturesFromBuffer } from 'ooxml-core/pptx';
import type { SignatureEntry, SignatureState, SignatureSummary } from 'ooxml-core/pptx';

import {
	BUILDING_BLOCK_OK,
	buildingBlockFailure,
	type BuildingBlockResult,
} from './building-block-actions';

/**
 * Framework-free digital-signature workflow of a host-owned shell. It reports the verified state
 * of the package's signatures (`SignatureSummary`): a deck is signed only when they verify.
 *
 * Honest limits: the browser verifies each signature (the signed parts' digests, the signature
 * value and the signer certificate's key) but has no trust store, so a valid signature is at best
 * `verified-untrusted`. The viewer cannot create a signature (there is no signing key or
 * certificate store here). Saving always writes the package without its signatures, so "remove"
 * means "the saved file will not carry them".
 */

/** One signature of the package, with its verified state. */
export type HostSignatureInfo = SignatureEntry;

export interface SignaturesState {
	/** The verified state of the package's signatures. */
	state: SignatureState;
	/** The deck is shown as signed: its signatures verified (`verified` or `verified-untrusted`). */
	signed: boolean;
	/** The package carries signature parts, whatever their state. */
	present: boolean;
	count: number;
	/** Saving drops the signatures (always true: edits invalidate them and the viewer cannot re-sign). */
	removedOnSave: true;
	/** What the host can ask for. Creating a signature and a trust decision are not available. */
	supported: { list: true; remove: true; create: false; trust: false };
}

export function describeSignaturesState(summary: SignatureSummary | undefined): SignaturesState {
	const state = summary?.state ?? 'unsigned';
	return {
		state,
		signed: state === 'verified' || state === 'verified-untrusted',
		present: state !== 'unsigned',
		count: summary?.signatures.length ?? 0,
		removedOnSave: true,
		supported: { list: true, remove: true, create: false, trust: false },
	};
}

/** List the signatures of a package's bytes (empty for an unsigned or unreadable package). */
export async function listPackageSignatures(
	content: ArrayBuffer | Uint8Array | null | undefined,
): Promise<HostSignatureInfo[]> {
	if (!content) {
		return [];
	}
	return (await inspectSignaturesFromBuffer(content)).signatures;
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
			if (!input.state.present) {
				return buildingBlockFailure('no-signatures');
			}
			input.markDirty();
			return input.complete(BUILDING_BLOCK_OK);
		},
		sign: () => buildingBlockFailure('unsupported'),
	};
}
