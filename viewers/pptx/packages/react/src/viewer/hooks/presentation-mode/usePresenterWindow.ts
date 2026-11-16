/**
 * usePresenterWindow: the audience window of a slide show, for the React
 * binding. The window, deck hand-over, BroadcastChannel protocol and Swap
 * Displays are the shared `createPresenterSessionController`, which every
 * binding uses; this hook ties it to the component lifecycle and publishes the
 * presenter console's snapshot while presenter view is up.
 */
import {
	createPresenterSessionController,
	PRESENTATION_CHANNEL_NAME,
	PRESENTATION_HASH,
	PRESENTATION_MESSAGE_ORIGIN,
	PRESENTATION_NONCE_KEY,
	parseAudienceNonce as parseSharedAudienceNonce,
} from 'ooxml-ui/pptx';
import type { PresentationSnapshot, PresenterSessionController } from 'ooxml-ui/pptx';
import { useCallback, useEffect, useRef, useState } from 'react';

import { useController } from '../useController';

/** BroadcastChannel name shared between presenter and audience tabs. */
export const PRESENTER_CHANNEL_NAME = PRESENTATION_CHANNEL_NAME;

/** Hash fragment used to identify the audience tab. */
export const AUDIENCE_HASH = PRESENTATION_HASH;

/** Unique origin identifier so we only react to our own messages. */
export const PRESENTER_MSG_ORIGIN = PRESENTATION_MESSAGE_ORIGIN;

/** Hash key used to pass the session nonce to the audience tab. */
export const AUDIENCE_NONCE_KEY = PRESENTATION_NONCE_KEY;

/** The session nonce of an audience tab (`#pptx-audience&nonce=<uuid>`), or null. */
export function parseAudienceNonce(): string | null {
	return parseSharedAudienceNonce();
}

export interface PresenterSlideChangeMessage {
	origin: typeof PRESENTER_MSG_ORIGIN;
	type: 'presenter-slide-change';
	slideIndex: number;
	sessionId: string;
}

export interface PresenterExitMessage {
	origin: typeof PRESENTER_MSG_ORIGIN;
	type: 'presenter-exit';
	sessionId: string;
}

export type PresenterMessage = PresenterSlideChangeMessage | PresenterExitMessage;

export function isPresenterMessage(data: unknown): data is PresenterMessage {
	if (typeof data !== 'object' || data === null) {
		return false;
	}
	const msg = data as Record<string, unknown>;
	return (
		msg.origin === PRESENTER_MSG_ORIGIN &&
		typeof msg.sessionId === 'string' &&
		(msg.type === 'presenter-slide-change' || msg.type === 'presenter-exit')
	);
}

/** Returns true if the current page was opened as an audience tab. */
export function isAudienceTab(): boolean {
	return window.location.hash.startsWith(AUDIENCE_HASH);
}

export interface UsePresenterWindowInput {
	currentSlideIndex: number;
	isPresenterMode: boolean;
	/** Raw PPTX bytes to share with the audience tab via IndexedDB. */
	content?: ArrayBuffer | Uint8Array | null;
	snapshot: PresentationSnapshot;
}

export interface UsePresenterWindowResult {
	openAudienceWindow: () => boolean;
	closeAudienceWindow: () => void;
	isAudienceWindowOpen: () => boolean;
	syncSlideToAudience: (slideIndex: number) => void;
	syncStateToAudience: (snapshot: PresentationSnapshot) => void;
	swapDisplays: () => Promise<boolean>;
}

export function usePresenterWindow(input: UsePresenterWindowInput): UsePresenterWindowResult {
	const { isPresenterMode, content, snapshot } = input;
	const contentRef = useRef(content);
	contentRef.current = content;
	const [session] = useState<PresenterSessionController>(() =>
		createPresenterSessionController({ getSource: () => contentRef.current }),
	);
	const { audienceOpen } = useController(session);

	useEffect(() => {
		if (isPresenterMode) {
			session.setSnapshot(snapshot);
		}
	}, [session, snapshot, isPresenterMode]);
	useEffect(() => {
		if (!isPresenterMode) {
			session.closeAudience();
		}
	}, [session, isPresenterMode]);
	useEffect(() => () => session.dispose(), [session]);

	return {
		openAudienceWindow: session.openAudience,
		closeAudienceWindow: session.closeAudience,
		isAudienceWindowOpen: useCallback(() => audienceOpen, [audienceOpen]),
		syncSlideToAudience: useCallback(
			(slideIndex: number) => session.updateSnapshot({ slideIndex }),
			[session],
		),
		syncStateToAudience: session.setSnapshot,
		swapDisplays: session.swapDisplays,
	};
}
