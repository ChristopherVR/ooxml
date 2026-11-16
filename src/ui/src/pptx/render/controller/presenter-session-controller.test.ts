import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPresenterSessionController } from './presenter-session-controller';
import type { PresenterSessionChannel } from './presenter-session-controller';

/** An in-memory BroadcastChannel: every channel from one bus hears the others. */
function createBus() {
	const channels = new Set<FakeChannel>();
	class FakeChannel implements PresenterSessionChannel {
		readonly sent: unknown[] = [];
		private readonly listeners = new Set<(event: MessageEvent) => void>();
		constructor() {
			channels.add(this);
		}
		postMessage(message: unknown): void {
			this.sent.push(message);
			for (const other of channels) {
				if (other !== this) {
					for (const listener of other.listeners) {
						listener({ data: message } as MessageEvent);
					}
				}
			}
		}
		addEventListener(_type: 'message', listener: (event: MessageEvent) => void): void {
			this.listeners.add(listener);
		}
		removeEventListener(_type: 'message', listener: (event: MessageEvent) => void): void {
			this.listeners.delete(listener);
		}
		close(): void {
			channels.delete(this);
		}
	}
	return () => new FakeChannel();
}

function fakeWindow(hash = '') {
	const popup = { closed: false, close: vi.fn(), location: { replace: vi.fn() } };
	const win = {
		location: { hash, href: `https://example.test/app${hash}` },
		open: vi.fn(() => popup),
		addEventListener: vi.fn(),
		removeEventListener: vi.fn(),
	};
	popup.close.mockImplementation(() => {
		popup.closed = true;
	});
	return { win: win as unknown as Window, popup, open: win.open };
}

afterEach(() => {
	vi.useRealTimers();
});

describe('createPresenterSessionController', () => {
	it('opens the audience window, stores the deck and navigates it', async () => {
		const { win, popup } = fakeWindow();
		const storeDeck = vi.fn(async () => undefined);
		const presenter = createPresenterSessionController({
			window: win,
			createChannel: createBus(),
			getSource: () => new Uint8Array([1]),
			storeDeck,
			clearDeck: async () => undefined,
		});
		expect(presenter.getSnapshot().role).toBe('presenter');
		expect(presenter.openAudience()).toBe(true);
		const { sessionId, audienceOpen } = presenter.getSnapshot();
		expect(audienceOpen).toBe(true);
		await vi.waitFor(() => expect(popup.location.replace).toHaveBeenCalled());
		expect(storeDeck).toHaveBeenCalledWith(sessionId, new Uint8Array([1]));
		expect(String(popup.location.replace.mock.calls[0]![0])).toContain(sessionId);
		presenter.dispose();
	});

	it('navigates without a deck and reports a blocked popup', async () => {
		const { win, popup, open } = fakeWindow();
		const presenter = createPresenterSessionController({
			window: win,
			createChannel: createBus(),
			clearDeck: async () => undefined,
		});
		presenter.openAudience();
		await vi.waitFor(() => expect(popup.location.replace).toHaveBeenCalled());
		open.mockReturnValueOnce(null as never);
		expect(presenter.openAudience()).toBe(false);
		expect(presenter.getSnapshot().audienceOpen).toBe(false);
	});

	it('keeps an audience in step and ends the session', () => {
		const bus = createBus();
		const host = fakeWindow();
		const clearDeck = vi.fn(async () => undefined);
		const onAudienceClosed = vi.fn();
		const presenter = createPresenterSessionController({
			window: host.win,
			createChannel: bus,
			clearDeck,
			onAudienceClosed,
		});
		presenter.openAudience();
		const { sessionId } = presenter.getSnapshot();
		const audienceWin = fakeWindow(`#pptx-audience&nonce=${sessionId}`).win;
		const onAudienceSlide = vi.fn();
		const onAudienceExit = vi.fn();
		presenter.updateSnapshot({ slideIndex: 2 });
		const audience = createPresenterSessionController({
			window: audienceWin,
			createChannel: bus,
			clearDeck,
			onAudienceSlide,
			onAudienceExit,
		});
		expect(audience.getSnapshot().role).toBe('audience');
		// Connecting announces "ready"; the presenter answers with the CURRENT snapshot.
		audience.connect();
		expect(onAudienceSlide).toHaveBeenLastCalledWith(2);
		presenter.updateSnapshot({ slideIndex: 4, blackout: 'black' });
		expect(audience.getSnapshot().snapshot).toMatchObject({ slideIndex: 4, blackout: 'black' });
		presenter.closeAudience();
		expect(onAudienceExit).toHaveBeenCalledTimes(1);
		expect(onAudienceClosed).toHaveBeenCalledTimes(1);
		expect(host.popup.close).toHaveBeenCalled();
		expect(clearDeck).toHaveBeenCalledWith(sessionId);
		expect(presenter.getSnapshot()).toMatchObject({ audienceOpen: false, sessionId: '' });
		audience.dispose();
		presenter.dispose();
	});

	it('ignores a different session and notices a popup the user closed', () => {
		vi.useFakeTimers();
		const bus = createBus();
		const host = fakeWindow();
		const presenter = createPresenterSessionController({
			window: host.win,
			createChannel: bus,
			clearDeck: async () => undefined,
		});
		presenter.openAudience();
		const onAudienceSlide = vi.fn();
		const stranger = createPresenterSessionController({
			window: fakeWindow('#pptx-audience&nonce=other').win,
			createChannel: bus,
			onAudienceSlide,
		});
		stranger.connect();
		presenter.updateSnapshot({ slideIndex: 3 });
		expect(onAudienceSlide).not.toHaveBeenCalled();
		host.popup.closed = true;
		vi.advanceTimersByTime(1000);
		expect(presenter.getSnapshot().audienceOpen).toBe(false);
		presenter.dispose();
		stranger.dispose();
	});

	it('does not publish while no audience window is open', () => {
		const bus = createBus();
		let channel: PresenterSessionChannel & { sent: unknown[] };
		const presenter = createPresenterSessionController({
			window: fakeWindow().win,
			createChannel: () => (channel = bus() as never),
		});
		presenter.connect();
		presenter.updateSnapshot({ slideIndex: 1 });
		expect(channel!.sent).toEqual([]);
		expect(presenter.getSnapshot().snapshot.slideIndex).toBe(1);
	});

	it('swapDisplays is false with no audience window', async () => {
		const presenter = createPresenterSessionController({ window: fakeWindow().win });
		await expect(presenter.swapDisplays()).resolves.toBe(false);
	});
});
