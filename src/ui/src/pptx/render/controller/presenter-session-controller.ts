/**
 * The presenter and audience windows of a slide show as a headless controller: opening the
 * audience window, handing it the deck, keeping it in step over the `BroadcastChannel`, the
 * Swap Displays command and ending the session, plus the audience side of the same protocol.
 *
 * A page opened with the `#pptx-audience&nonce=...` hash is the audience; any other page is a
 * presenter. The controller owns the {@link PresentationSnapshot} both sides share: a presenter
 * publishes it on every change while an audience window is open, an audience adopts each one it
 * receives.
 *
 * One behaviour where the bindings had drifted (see `docs/structure-plan.md` section 5):
 *
 *  - the popup opens synchronously inside the user gesture, then the deck is stored and the
 *    popup navigated; with no deck to hand over it is navigated anyway (the host may load its
 *    own source there), and a failed hand-over closes it;
 *  - "audience ready" is answered with the current snapshot, never one captured at open;
 *  - a popup the user closes is noticed within a second and the session ended;
 *  - the stored deck is cleared on close, on page unload and when the audience receives the
 *    exit message.
 *
 * @module render/controller/presenter-session-controller
 */

import { mergePresentationSnapshot } from '../presenter-console';
import {
	buildPresentationAudienceUrl,
	clearPresentationDeck,
	createInitialPresentationSnapshot,
	createPresentationSessionId,
	isPresentationSessionMessage,
	parsePresentationSessionId,
	placeAudienceWindow,
	PRESENTATION_CHANNEL_NAME,
	PRESENTATION_MESSAGE_ORIGIN,
	resolveAudienceScreenPlacement,
	storePresentationDeck,
	swapPresentationWindows,
} from '../presentation-session';
import type { PresentationSnapshot } from '../presentation-session-types';
import { createControllerStore } from './controller';
import type { PptxController } from './controller';

export type PresenterSessionRole = 'presenter' | 'audience';

export interface PresenterSessionState {
	readonly role: PresenterSessionRole;
	/** A presenter's audience window is open. Always false for an audience. */
	readonly audienceOpen: boolean;
	/** The open session's id, or '' when none. */
	readonly sessionId: string;
	/** The state both windows share. */
	readonly snapshot: PresentationSnapshot;
}

/** The part of `BroadcastChannel` the controller uses. */
export interface PresenterSessionChannel {
	postMessage(message: unknown): void;
	addEventListener(type: 'message', listener: (event: MessageEvent) => void): void;
	removeEventListener(type: 'message', listener: (event: MessageEvent) => void): void;
	close(): void;
}

type DeckBytes = ArrayBuffer | Uint8Array;

export interface PresenterSessionControllerDeps {
	/** The page's window; defaults to the global `window` (none during SSR). */
	window?: Window | null;
	/** The deck to hand the audience window, or null when there is none. May be async. */
	getSource?(): DeckBytes | null | undefined | Promise<DeckBytes | null | undefined>;
	/** Audience: the presenter's snapshot arrived (after the state was adopted). */
	onAudienceSnapshot?(snapshot: PresentationSnapshot): void;
	/** Audience: the presenter moved to a slide (a snapshot or a bare slide change). */
	onAudienceSlide?(slideIndex: number): void;
	/** Audience: the presenter ended the session. The host decides what that looks like. */
	onAudienceExit?(): void;
	/** Presenter: the audience window was closed, by the user or by {@link PresenterSessionController.closeAudience}. */
	onAudienceClosed?(): void;
	/** The first snapshot's slide; defaults to 0. */
	initialSlideIndex?: number;
	createChannel?(): PresenterSessionChannel | null;
	storeDeck?(sessionId: string, bytes: DeckBytes): Promise<void>;
	clearDeck?(sessionId: string): Promise<void>;
	/** How often a presenter checks that the popup is still open; defaults to 1000 ms. */
	pollMs?: number;
}

export interface PresenterSessionController extends PptxController<PresenterSessionState> {
	/** Start listening (idempotent); an audience announces itself to its presenter. */
	connect(): void;
	/** Presenter: open the audience window. False when the popup was blocked. */
	openAudience(): boolean;
	/** Presenter: tell the audience the session ended and close its window. */
	closeAudience(): void;
	/** Merge `patch` into the shared snapshot and publish it. */
	updateSnapshot(patch: Partial<PresentationSnapshot>): void;
	/** Replace the shared snapshot (a binding that owns the snapshot elsewhere) and publish it. */
	setSnapshot(snapshot: PresentationSnapshot): void;
	/** Presenter: PowerPoint's Swap Displays. False when no audience window or no Window Management API. */
	swapDisplays(): Promise<boolean>;
	/** The open audience window, for hosts that position or focus it. */
	audienceWindow(): Window | null;
	/** Close the audience window (presenter) and the channel. */
	dispose(): void;
}

function defaultChannel(): PresenterSessionChannel | null {
	try {
		return typeof BroadcastChannel === 'undefined'
			? null
			: new BroadcastChannel(PRESENTATION_CHANNEL_NAME);
	} catch {
		return null;
	}
}

function hostWindow(deps: PresenterSessionControllerDeps): Window | null {
	if (deps.window !== undefined) {
		return deps.window;
	}
	return typeof window === 'undefined' ? null : window;
}

/** Create the presenter-session controller for this page. */
export function createPresenterSessionController(
	deps: PresenterSessionControllerDeps = {},
): PresenterSessionController {
	const win = hostWindow(deps);
	const audienceSessionId = win ? parsePresentationSessionId(win.location.hash) : null;
	const role: PresenterSessionRole = audienceSessionId === null ? 'presenter' : 'audience';
	const storeDeck = deps.storeDeck ?? storePresentationDeck;
	const clearDeck = deps.clearDeck ?? clearPresentationDeck;
	const pollMs = deps.pollMs ?? 1000;
	const store = createControllerStore<PresenterSessionState>({
		role,
		audienceOpen: false,
		sessionId: audienceSessionId ?? '',
		snapshot: createInitialPresentationSnapshot(deps.initialSlideIndex ?? 0),
	});

	let channel: PresenterSessionChannel | null | undefined;
	let popup: Window | null = null;
	let poll: ReturnType<typeof setInterval> | null = null;
	let connected = false;

	const getChannel = (): PresenterSessionChannel | null => {
		if (channel === undefined) {
			channel = (deps.createChannel ?? defaultChannel)();
		}
		return channel;
	};
	const post = (message: Record<string, unknown>): void => {
		try {
			getChannel()?.postMessage({ origin: PRESENTATION_MESSAGE_ORIGIN, ...message });
		} catch {
			// A closed channel: nothing to tell.
		}
	};
	const publish = (): void => {
		const { role: current, sessionId, snapshot } = store.getSnapshot();
		if (current === 'presenter' && sessionId) {
			post({ type: 'presenter-state', sessionId, snapshot });
		}
	};
	const stopPoll = (): void => {
		if (poll !== null) {
			clearInterval(poll);
			poll = null;
		}
	};
	const endSession = (closeWindow: boolean): void => {
		const { sessionId, audienceOpen } = store.getSnapshot();
		if (role !== 'presenter' || (!sessionId && !popup)) {
			return;
		}
		if (sessionId) {
			post({ type: 'presenter-exit', sessionId });
			void clearDeck(sessionId);
		}
		if (closeWindow) {
			try {
				popup?.close();
			} catch {
				// A window we can no longer reach.
			}
		}
		popup = null;
		stopPoll();
		store.update((state) => ({ ...state, audienceOpen: false, sessionId: '' }));
		if (audienceOpen) {
			deps.onAudienceClosed?.();
		}
	};

	const onMessage = (event: MessageEvent): void => {
		const message: unknown = event.data;
		if (!isPresentationSessionMessage(message)) {
			return;
		}
		const state = store.getSnapshot();
		if (role === 'audience') {
			if (message.sessionId !== audienceSessionId) {
				return;
			}
			if (message.type === 'presenter-state') {
				store.set({ ...state, snapshot: message.snapshot });
				deps.onAudienceSnapshot?.(message.snapshot);
				deps.onAudienceSlide?.(message.snapshot.slideIndex);
			} else if (message.type === 'presenter-slide-change') {
				deps.onAudienceSlide?.(message.slideIndex);
			} else if (message.type === 'presenter-exit') {
				void clearDeck(message.sessionId);
				deps.onAudienceExit?.();
			}
		} else if (message.type === 'audience-ready' && message.sessionId === state.sessionId) {
			publish();
		}
	};
	const onUnload = (): void => {
		const { sessionId } = store.getSnapshot();
		if (sessionId && role === 'presenter') {
			void clearDeck(sessionId);
		}
	};

	const connect = (): void => {
		if (connected) {
			return;
		}
		connected = true;
		getChannel()?.addEventListener('message', onMessage);
		if (role === 'audience' && audienceSessionId) {
			post({ type: 'audience-ready', sessionId: audienceSessionId });
		} else {
			win?.addEventListener('beforeunload', onUnload);
		}
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		connect,
		openAudience() {
			if (role !== 'presenter' || !win) {
				return false;
			}
			endSession(true);
			connect();
			// Opened synchronously inside the user gesture so a popup blocker lets it through;
			// the deck is stored while the blank window waits.
			const opened = win.open(
				'about:blank',
				'pptx-viewer-audience',
				'popup=yes,width=1280,height=720',
			);
			if (!opened) {
				return false;
			}
			popup = opened;
			const sessionId = createPresentationSessionId();
			store.update((state) => ({ ...state, audienceOpen: true, sessionId }));
			const url = buildPresentationAudienceUrl(win.location.href, sessionId);
			void resolveAudienceScreenPlacement(win).then((placement) => {
				if (placement && popup === opened && !opened.closed) {
					placeAudienceWindow(opened, placement);
				}
				return undefined;
			});
			const current = (): boolean =>
				popup === opened && !opened.closed && store.getSnapshot().sessionId === sessionId;
			void Promise.resolve()
				.then(() => deps.getSource?.())
				.then(async (bytes) => {
					if (bytes && current()) {
						await storeDeck(sessionId, bytes);
					}
					if (current()) {
						opened.location.replace(url);
					}
					return undefined;
				})
				.catch(() => {
					if (current()) {
						endSession(true);
					}
				});
			poll = setInterval(() => {
				if (popup === opened && opened.closed) {
					endSession(false);
				}
			}, pollMs);
			return true;
		},
		closeAudience: () => endSession(true),
		updateSnapshot(patch) {
			const state = store.getSnapshot();
			const snapshot = mergePresentationSnapshot(state.snapshot, patch);
			if (snapshot !== state.snapshot) {
				store.set({ ...state, snapshot });
				publish();
			}
		},
		setSnapshot(snapshot) {
			const state = store.getSnapshot();
			if (snapshot !== state.snapshot) {
				store.set({ ...state, snapshot });
				publish();
			}
		},
		async swapDisplays() {
			const target = popup;
			if (!win || !target || target.closed) {
				return false;
			}
			return swapPresentationWindows(win, target);
		},
		audienceWindow: () => (popup && !popup.closed ? popup : null),
		dispose() {
			endSession(true);
			win?.removeEventListener('beforeunload', onUnload);
			if (channel) {
				channel.removeEventListener('message', onMessage);
				try {
					channel.close();
				} catch {
					// Already closed.
				}
			}
			channel = undefined;
			connected = false;
		},
	};
}
