import type {
	CollabLoadOrigin,
	CollaborationHistoryTarget,
	CollaborationLivePatcher,
	DeckSaveState,
} from 'ooxml-ui/pptx';
import {
	bindCollaborationHistory,
	buildDeckSaveOptions,
	collaborationHistoryReady,
	readSlidesFromYDoc,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { useCallback, useEffect, useRef } from 'react';

import { collabModules } from './collaboration-modules';
import type { CollaborationContextValue } from './types';
import { useCollaborationLivePatch } from './useCollaborationLivePatch';
import { useYjsDocumentSync } from './useYjsDocumentSync';

export interface CollaborationDocumentSyncInput {
	collaboration: CollaborationContextValue | null;
	/** The editor history; Undo and Redo go through the room's own history while it is writable. */
	history?: CollaborationHistoryTarget;
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
	setSlides: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
	content: ArrayBuffer | Uint8Array | null;
	loadVersion: number;
	loadOrigin: CollabLoadOrigin;
	livePatcher: CollaborationLivePatcher;
	deckSaveState: DeckSaveState;
	onReadOnlyChange?: (readOnly: boolean) => void;
}

/** Shared wiring for the bundled editor and custom toolbar/canvas hosts. */
export function useCollaborationDocumentSync(input: CollaborationDocumentSyncInput): void {
	const { collaboration, history, ...state } = input;
	const latest = useRef(input);
	latest.current = input;
	const getSourceBytes = useCallback(() => {
		const bytes = latest.current.content;
		return bytes instanceof Uint8Array ? bytes : bytes ? new Uint8Array(bytes) : null;
	}, []);
	const getSaveOptions = useCallback(() => buildDeckSaveOptions(latest.current.deckSaveState), []);
	const config = collaboration?.config;
	const doc = collaboration?.doc;
	const status = collaboration?.status;
	const preserveAcceptedEdits = useCallback((detachingDoc: NonNullable<typeof doc>) => {
		const current = latest.current;
		const session = current.collaboration;
		if (
			session?.doc === detachingDoc &&
			!session.config.externalSession &&
			(session.status !== 'connected' || !session.synced || session.config.role === 'viewer')
		) {
			current.setSlides(readSlidesFromYDoc(detachingDoc));
		}
	}, []);
	// The host can keep a synchronized document editable while offline.
	const isConnected = config?.externalSession
		? Boolean(collaboration?.doc)
		: collaboration?.status === 'connected';
	const historyReady = collaborationHistoryReady({
		hasDoc: Boolean(doc),
		external: Boolean(config?.externalSession),
		connected: isConnected && !config?.externalSession,
		synced: collaboration?.synced ?? false,
		viewer: config?.role === 'viewer',
	});
	// The history object is rebuilt every render; its attach function is the stable identity.
	const attachCollaboration = history?.attachCollaboration;
	useEffect(() => {
		if (!attachCollaboration || !doc || !historyReady) {
			return;
		}
		return bindCollaborationHistory({ attachCollaboration }, doc, collabModules.loadYjs);
	}, [attachCollaboration, doc, historyReady]);
	useYjsDocumentSync({
		...state,
		doc: collaboration?.doc ?? null,
		isConnected,
		isSynced: collaboration?.synced ?? false,
		config,
		livePatcher: state.livePatcher,
		getSourceBytes,
		getSaveOptions,
		...(collaboration?.reportWriteBack ? { onWriteBackStatus: collaboration.reportWriteBack } : {}),
	});
	const initialized = useCollaborationLivePatch({
		patcher: state.livePatcher,
		externalSession: config?.externalSession,
		doc: collaboration?.doc ?? null,
		isConnected: isConnected && config?.role !== 'viewer',
		isSynced: collaboration?.synced ?? false,
		onBeforeDetach: preserveAcceptedEdits,
	});
	useEffect(() => {
		if (!config?.externalSession || !doc) {
			latest.current.onReadOnlyChange?.(
				config?.externalSession
					? status !== 'error'
					: Boolean(config) &&
							status !== 'error' &&
							(!initialized || (Boolean(doc) && config?.role === 'viewer')),
			);
		}
	}, [doc, config, status, initialized]);
}
