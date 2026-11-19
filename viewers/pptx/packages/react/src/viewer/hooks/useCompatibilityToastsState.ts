import type { UpdateRefusalCode } from 'ooxml-core/collab';
import type { CompatibilityWarningToast } from 'ooxml-ui/pptx';
import { collaborationRefusalToast, createRemoteValueNotifier } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { useCallback, useEffect, useState } from 'react';

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
