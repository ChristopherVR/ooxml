/**
 * presenter-window.service.ts: the presenter and audience windows for the
 * Angular binding, an adapter over the shared `createPresenterSessionController`
 * (the window, deck hand-over, BroadcastChannel protocol and Swap Displays),
 * which every binding uses, so a presenter from any framework drives an
 * audience tab rendered by any other.
 */

import { Injectable } from '@angular/core';
import {
	createPresenterSessionController,
	PRESENTATION_CHANNEL_NAME,
	PRESENTATION_MESSAGE_ORIGIN,
	PRESENTATION_NONCE_KEY,
} from 'ooxml-ui/pptx';
import type { PresentationSnapshot, PresenterSessionController } from 'ooxml-ui/pptx';

import { controllerSignal } from './controller-signal';

export { parseAudienceNonce } from 'ooxml-ui/pptx';

/** BroadcastChannel name shared between presenter and audience tabs. */
export const PRESENTER_CHANNEL_NAME = PRESENTATION_CHANNEL_NAME;

/** Unique origin identifier so we only react to our own messages. */
export const PRESENTER_MSG_ORIGIN = PRESENTATION_MESSAGE_ORIGIN;

/** Hash key used to pass the session nonce to the audience tab. */
export const AUDIENCE_NONCE_KEY = PRESENTATION_NONCE_KEY;

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

/** Type guard for messages arriving on the presenter BroadcastChannel. */
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

@Injectable({ providedIn: 'root' })
export class PresenterWindowService {
	private content: ArrayBuffer | Uint8Array | null = null;
	private onSlide: (index: number) => void = () => undefined;
	private onExit: () => void = () => undefined;
	private readonly session: PresenterSessionController = createPresenterSessionController({
		getSource: () => this.content,
		onAudienceSlide: (index) => this.onSlide(index),
		onAudienceExit: () => this.onExit(),
	});
	private readonly state = controllerSignal(this.session);

	/** The snapshot both windows share (reactive). */
	snapshot(): PresentationSnapshot {
		return this.state().snapshot;
	}

	/** Whether the audience window is open (reactive). */
	isAudienceWindowOpen(): boolean {
		return this.state().audienceOpen;
	}

	/** PowerPoint's "Swap Displays". False with no audience window or no Window Management API. */
	swapDisplays(): Promise<boolean> {
		return this.session.swapDisplays();
	}

	syncSlideToAudience(slideIndex: number): void {
		this.updateSnapshot({ slideIndex });
	}

	updateSnapshot(patch: Partial<PresentationSnapshot>): void {
		this.session.updateSnapshot(patch);
	}

	closeAudienceWindow(): void {
		this.session.closeAudience();
	}

	/** Open the audience tab and hand it `content`. False when the popup is blocked. */
	openAudienceWindow(content: ArrayBuffer | Uint8Array | null, currentSlideIndex: number): boolean {
		this.content = content;
		this.session.updateSnapshot({ slideIndex: currentSlideIndex });
		return this.session.openAudience();
	}

	/** Connect an audience tab to the presenter channel and announce readiness. */
	connectAudience(onSlide: (index: number) => void, onExit: () => void): () => void {
		if (this.session.getSnapshot().role !== 'audience') {
			return () => undefined;
		}
		this.onSlide = onSlide;
		this.onExit = onExit;
		this.session.connect();
		return () => this.session.dispose();
	}
}
