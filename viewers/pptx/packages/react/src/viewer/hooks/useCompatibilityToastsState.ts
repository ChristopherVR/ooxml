import type { UpdateRefusalCode } from 'ooxml-core/collab';
import type { CompatibilityWarningToast } from 'ooxml-ui/pptx';
import {
	collaborationRefusalToast,
	createRemoteValueNotifier,
	signatureLoadNotice,
} from 'ooxml-ui/pptx';
import type { PptxSlide, SignatureSummary } from 'pptx-viewer-core';
import { useCallback, useEffect, useMemo, useState } from 'react';

import { useSignatureStatusView } from '../components/signature-look';

/**
 * useCompatibilityToastsState: the load-diagnostic toast stack for
 * `getCompatibilityWarnings()`-shaped deck/slide warnings.
 *
 * `setToasts` is threaded into `useLoadContent` exactly like every other
 * per-load setter: each load overwrites the array wholesale, which is also
 * how dismissal state naturally resets on the next load without a separate
 * reset effect (a fresh empty or non-empty array replaces whatever was
 * dismissed from the previous deck).
 */
export function useCompatibilityToastsState() {
	const [toasts, setToasts] = useState<CompatibilityWarningToast[]>([]);

	const dismiss = useCallback((id: string) => {
		setToasts((prev) => prev.filter((toast) => toast.id !== id));
	}, []);

	const dismissAll = useCallback(() => {
		setToasts([]);
	}, []);

	return { toasts, setToasts, dismiss, dismissAll };
}

export type UseCompatibilityToastsStateResult = ReturnType<typeof useCompatibilityToastsState>;

/**
 * Add the toast for values a collaboration peer sent that were left out,
 * whenever the slides gain such a rejection (however they were adopted from
 * the room). A rejection already reported does not bring a dismissed toast back.
 */
export function useRemoteValueToast(
	slides: readonly PptxSlide[],
	setToasts: UseCompatibilityToastsStateResult['setToasts'],
): void {
	const [notice] = useState(createRemoteValueNotifier);
	useEffect(() => {
		const toast = notice(slides);
		if (toast) {
			setToasts((prev) => [toast, ...prev.filter((entry) => entry.id !== toast.id)]);
		}
	}, [notice, slides, setToasts]);
}

/**
 * Add the toast that says why sharing stopped whenever the collaboration
 * session reports a refused update (it replaces an earlier one).
 */
export function useCollaborationRefusalToast(
	refusal: UpdateRefusalCode | null | undefined,
	setToasts: UseCompatibilityToastsStateResult['setToasts'],
): void {
	useEffect(() => {
		if (!refusal) {
			return;
		}
		const toast = collaborationRefusalToast(refusal);
		setToasts((prev) => [toast, ...prev.filter((entry) => entry.id !== toast.id)]);
	}, [refusal, setToasts]);
}

/**
 * The toast stack with the signature load notice (the shared `signatureLoadNotice`) on top when
 * the loaded deck carries signatures. The notice follows the UI language; dismissing it (or the
 * whole stack) hides it until another deck loads, since every load sets a new summary.
 */
export function useSignatureNotice(
	signatures: SignatureSummary | undefined,
	stack: UseCompatibilityToastsStateResult,
): Pick<UseCompatibilityToastsStateResult, 'toasts' | 'dismiss' | 'dismissAll'> {
	const view = useSignatureStatusView(signatures);
	const [dismissedFor, setDismissedFor] = useState<SignatureSummary | undefined>();
	const notice = dismissedFor === signatures ? undefined : signatureLoadNotice(view);
	const { toasts: stored, dismiss: dismissStored, dismissAll: dismissAllStored } = stack;
	const toasts = useMemo(() => (notice ? [notice, ...stored] : stored), [notice, stored]);
	const dismiss = useCallback(
		(id: string) => {
			if (id === notice?.id) setDismissedFor(signatures);
			else dismissStored(id);
		},
		[notice, signatures, dismissStored],
	);
	const dismissAll = useCallback(() => {
		setDismissedFor(signatures);
		dismissAllStored();
	}, [signatures, dismissAllStored]);
	return { toasts, dismiss, dismissAll };
}
