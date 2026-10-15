import { type ActivityItem, type NotificationSettings, describeActivity } from 'ooxml-core/teams';

/** The browser features delivery uses, injectable so the rules are testable without a browser. */
export interface DeliveryEnv {
	/** The tab is in the background (the Page Visibility API). */
	hidden: () => boolean;
	/** The Notification API permission, or `unsupported`. */
	permission: () => NotificationPermission | 'unsupported';
	/** Show a desktop notification; `onClick` runs when it is clicked. */
	show: (title: string, options: { body: string; tag: string }, onClick: () => void) => void;
	beep: () => void;
}

export interface DeliveryActions {
	/** An in-app toast for the item. */
	toast: (item: ActivityItem) => void;
	/** Open the item (what clicking a notification does). */
	open: (item: ActivityItem) => void;
}

/**
 * Announce an item the core decided to notify about (level and Do not disturb are already
 * applied there). A visible tab gets an in-app toast; a hidden tab gets a desktop notification
 * when permission was granted. The sound plays in both cases when it is on.
 */
export function deliverNotification(
	item: ActivityItem,
	settings: Pick<NotificationSettings, 'sound'>,
	env: DeliveryEnv,
	actions: DeliveryActions,
): 'toast' | 'desktop' | 'none' {
	let result: 'toast' | 'desktop' | 'none' = 'none';
	if (!env.hidden()) {
		actions.toast(item);
		result = 'toast';
	} else if (env.permission() === 'granted') {
		env.show(describeActivity(item), { body: item.preview, tag: item.id }, () =>
			actions.open(item),
		);
		result = 'desktop';
	}
	if (result !== 'none' && settings.sound) env.beep();
	return result;
}

/** A short two-note chime from the Web Audio API; silent where audio is unavailable or locked. */
function chime(): void {
	try {
		const Context =
			globalThis.AudioContext ??
			(globalThis as unknown as { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
		if (!Context) return;
		const audio = new Context();
		const gain = audio.createGain();
		gain.gain.setValueAtTime(0.06, audio.currentTime);
		gain.gain.exponentialRampToValueAtTime(0.0001, audio.currentTime + 0.35);
		gain.connect(audio.destination);
		[660, 880].forEach((frequency, index) => {
			const tone = audio.createOscillator();
			tone.frequency.value = frequency;
			tone.connect(gain);
			tone.start(audio.currentTime + index * 0.12);
			tone.stop(audio.currentTime + index * 0.12 + 0.12);
		});
		globalThis.setTimeout(() => void audio.close(), 700);
	} catch {
		/* Audio is optional. */
	}
}

/** The real browser, with every access guarded. */
export function browserDeliveryEnv(): DeliveryEnv {
	return {
		hidden: () => globalThis.document?.visibilityState === 'hidden',
		permission: () => globalThis.Notification?.permission ?? 'unsupported',
		show(title, options, onClick) {
			try {
				const notification = new globalThis.Notification(title, options);
				notification.onclick = () => {
					globalThis.focus();
					notification.close();
					onClick();
				};
			} catch {
				/* Some browsers only allow notifications from a service worker. */
			}
		},
		beep: chime,
	};
}
