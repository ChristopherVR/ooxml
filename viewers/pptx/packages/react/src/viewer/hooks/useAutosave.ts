import {
	createAutosaveController,
	resolveAutosaveIntervalMs,
	saveAutosaveSnapshot,
} from 'ooxml-ui/pptx';
import type { DocumentStorage, InlineTextDraftSink, PptxAutosaveStatus } from 'ooxml-ui/pptx';
import { useEffect, useRef, useState } from 'react';
import type { MutableRefObject } from 'react';

import { useController } from './useController';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** The one autosave status every binding renders (`PptxAutosaveStatus` in `ooxml-ui/pptx`). */
export type AutosaveStatus = PptxAutosaveStatus;

export interface UseAutosaveInput {
	/** Whether the document has unsaved changes. */
	isDirty: boolean;
	/** File path or name of the currently-open PPTX. Required for autosave to work. */
	filePath: string | undefined;
	/** Serialise current editor state to a Uint8Array. */
	serializeSlides: () => Promise<Uint8Array | null>;
	/**
	 * Autosave interval in seconds.
	 *
	 * @deprecated Pass {@link intervalMs} instead. Seconds cannot express the
	 * short cadence a host may legitimately ask for. Ignored when `intervalMs`
	 * is given.
	 */
	intervalSeconds?: number;
	/**
	 * Autosave interval in milliseconds, already resolved by the shared
	 * `resolveAutosaveIntervalMs` (host prop > File > Options > Save AutoRecover
	 * cadence > 120s default). Wins over {@link intervalSeconds}.
	 */
	intervalMs?: number;
	/**
	 * Whether autosave is enabled: the resolved `active` from the shared
	 * `resolveAutosaveActivation`, NOT the raw title-bar toggle.
	 */
	enabled?: boolean;
	/** Host storage; its recovery methods replace IndexedDB when it has them. */
	storage?: DocumentStorage;
	/**
	 * The values a snapshot is built from (reassigned, never mutated), read on
	 * every render. While `isDirty`, a change in any of them is an edit that
	 * starts the AutoRecover interval; unchanged values never cause a rewrite.
	 * Omit it and only the clean-to-dirty transition counts as an edit.
	 */
	getChangeSources?: () => readonly unknown[];
	/** Called with the serialised deck after each snapshot is written (the host's `onAutosave`). */
	onSaved?: ((data: Uint8Array) => void) | undefined;
	/** Receives the engine so in-progress inline text counts as an edit (`reportInlineTextDraft`). */
	inlineDraftSinkRef?: MutableRefObject<InlineTextDraftSink | null>;
	/** The element whose inline editor is open; a closed editor's draft is dropped. */
	getInlineEditingId?: () => string | null;
}

export interface UseAutosaveResult {
	/** Current autosave status for display in the StatusBar. */
	autosaveStatus: AutosaveStatus;
	/** Manually trigger an autosave right now. */
	triggerAutosave: () => Promise<void>;
}

// ---------------------------------------------------------------------------
// Hook: the React adapter of the shared autosave controller
// ---------------------------------------------------------------------------

export function useAutosave(input: UseAutosaveInput): UseAutosaveResult {
	const { isDirty, filePath, enabled = true } = input;
	const intervalMs = resolveAutosaveIntervalMs({
		hostIntervalMs: input.intervalMs,
		optionsIntervalSeconds: input.intervalSeconds,
	});

	// The controller reads the latest props through this ref, so it is built once.
	const live = useRef(input);
	live.current = input;
	const [controller] = useState(() =>
		createAutosaveController({
			getActivation: () => ({ active: live.current.enabled ?? true }),
			getFilePath: () => live.current.filePath,
			getIntervalMs: () =>
				resolveAutosaveIntervalMs({
					hostIntervalMs: live.current.intervalMs,
					optionsIntervalSeconds: live.current.intervalSeconds,
				}),
			serialize: () => live.current.serializeSlides(),
			getStorage: () => live.current.storage,
			onSaved: (data) => live.current.onSaved?.(data),
			writeSnapshot: saveAutosaveSnapshot,
			// serializeSlides already overlays the live inline draft from its refs.
			getInlineEditingId: () => live.current.getInlineEditingId?.(),
		}),
	);
	if (input.inlineDraftSinkRef) {
		input.inlineDraftSinkRef.current = controller;
	}
	const { status } = useController(controller);

	useEffect(() => {
		controller.refresh();
	}, [controller, enabled, filePath, intervalMs]);

	// Every render: the controller compares identities, so this is cheap, and it
	// also sees values that live in refs (an inline edit in progress).
	useEffect(() => {
		controller.observe(isDirty, live.current.getChangeSources?.() ?? []);
	});

	// No permanent teardown: StrictMode re-runs effects on the same controller.
	useEffect(() => () => controller.markClean(), [controller]);

	return {
		autosaveStatus: status,
		triggerAutosave: () => controller.saveNow(),
	};
}
