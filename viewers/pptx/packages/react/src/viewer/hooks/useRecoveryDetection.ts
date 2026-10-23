/**
 * useRecoveryDetection: the React adapter of the shared crash-recovery offer
 * (`createAutosaveRecoveryController` in `ooxml-ui/pptx`). After a deck
 * finishes loading it asks the recovery store whether a snapshot exists for it
 * and hands back the prompt descriptor the native dialog renders.
 */
import {
	createAutosaveRecoveryController,
	discardAutosaveRecovery,
	probeAutosaveRecovery,
} from 'ooxml-ui/pptx';
import type { AutosaveRecord, AutosaveRecoveryPrompt, DocumentStorage } from 'ooxml-ui/pptx';
import { useEffect, useRef, useState } from 'react';

import { useController } from './useController';

// ---------------------------------------------------------------------------
// Input / output
// ---------------------------------------------------------------------------

export interface UseRecoveryDetectionInput {
	filePath: string | undefined;
	/** Public document name rendered in the recovery message. */
	fileName?: string;
	loading: boolean;
	error: string | null;
	slideCount: number;
	/**
	 * Whether the host permits autosave (the shared `resolveAutosaveActivation`
	 * ceiling). A host that passed `autosave={false}` is never offered snapshots
	 * an earlier configuration left behind. Defaults to true.
	 */
	autosaveAllowed?: boolean;
	/** Host storage; the snapshot is read from its recovery methods when it has them. */
	storage?: DocumentStorage;
	/** Called with the stored record when a recovery snapshot is offered (the host's `onAutosaveRecovery`). */
	onOffered?: ((record: AutosaveRecord) => void) | undefined;
	/** Load the recovered bytes into the viewer. */
	onRestore?: (bytes: Uint8Array) => void;
	/** Legacy behaviour: also open the Version History panel when one is found. */
	openVersionHistory?: () => void;
}

export interface UseRecoveryDetectionResult {
	/** What the dialog should say, or null when there is nothing to offer. */
	prompt: AutosaveRecoveryPrompt | null;
	/** Accept: load the snapshot and close. */
	restore: () => void;
	/** Decline: drop the snapshot and close. */
	discard: () => void;
	/** The destructive delete is still pending. */
	discarding: boolean;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useRecoveryDetection(input: UseRecoveryDetectionInput): UseRecoveryDetectionResult {
	const { filePath, fileName, loading, error, slideCount, autosaveAllowed = true } = input;
	const live = useRef(input);
	live.current = input;
	const [controller] = useState(() =>
		createAutosaveRecoveryController({
			getStorage: () => live.current.storage,
			load: (bytes) => live.current.onRestore?.(bytes),
			onOffered: (record) => {
				live.current.onOffered?.(record);
				live.current.openVersionHistory?.();
			},
			probe: probeAutosaveRecovery,
			discardRecord: discardAutosaveRecovery,
		}),
	);
	const { prompt, discarding } = useController(controller);

	useEffect(() => {
		controller.check({ filePath, fileName, loading, error, slideCount, autosaveAllowed });
	}, [controller, filePath, fileName, loading, error, slideCount, autosaveAllowed]);

	return {
		prompt,
		discarding,
		restore: () => {
			void controller.restore().catch(() => {
				// The load path reports its own error.
			});
		},
		discard: () => {
			void controller.discard();
		},
	};
}
