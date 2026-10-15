import { html } from 'lit';
import type { NotificationLevel, NotificationSettings } from 'ooxml-core/teams';
import type { OfficeUiSwitch } from '../../form/switch';

const LEVELS: readonly (readonly [NotificationLevel, string, string])[] = [
	['all', 'All activity', 'Mentions, replies, reactions and pinned messages'],
	['mentions', 'Mentions only', 'Only when someone mentions you'],
	['off', 'Off', 'Activity still collects in the Activity feed'],
];

export type DesktopPermission = NotificationPermission | 'unsupported';

const PERMISSION_TEXT: Record<DesktopPermission, string> = {
	granted:
		'Desktop notifications are on for this browser. They appear when this tab is in the background.',
	default:
		'Desktop notifications appear when this tab is in the background. Your browser asks first.',
	denied:
		'Desktop notifications are blocked for this site. Allow them in your browser settings to turn them on.',
	unsupported: 'This browser does not support desktop notifications.',
};

/** What this browser currently allows; safe where the Notification API is missing. */
export const desktopPermission = (): DesktopPermission =>
	globalThis.Notification?.permission ?? 'unsupported';

/** Level, sound and desktop permission controls. Native radios keep keyboard and screen readers. */
export function notificationChoices(
	settings: NotificationSettings,
	permission: DesktopPermission,
	handlers: {
		change: (next: Partial<NotificationSettings>) => void;
		requestPermission: () => void;
	},
) {
	return html`<fieldset class="notification-level">
			<legend>Notify me about</legend>
			${LEVELS.map(
				([level, label, description]) =>
					html`<label class="notification-choice">
						<input
							type="radio"
							name="notification-level"
							value=${level}
							.checked=${settings.level === level}
							@change=${() => handlers.change({ level })}
						/>
						<span><strong>${label}</strong><small>${description}</small></span>
					</label>`,
			)}
		</fieldset>
		<h3>Sound</h3>
		<label class="preference"
			><span>Play a sound for new notifications</span
			><office-ui-switch
				aria-label="Notification sound"
				.checked=${settings.sound}
				@change=${(event: Event) => handlers.change({ sound: (event.target as OfficeUiSwitch).checked })}
			></office-ui-switch
		></label>
		<h3>Desktop notifications</h3>
		<p class="description" role="status">${PERMISSION_TEXT[permission]}</p>
		${
			permission === 'default'
				? html`<button type="button" class="primary" @click=${handlers.requestPermission}>
						Enable desktop notifications
					</button>`
				: ''
		}
		<p class="description">
			Do not disturb silences every notification. Email and mobile push notifications are not
			available.
		</p>`;
}
