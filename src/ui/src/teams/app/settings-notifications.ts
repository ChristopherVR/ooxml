import { html } from 'lit';
import type { NotificationLevel, NotificationSettings } from 'ooxml-core/teams';
import type { OfficeUiSwitch } from '../../form/switch';
import { teamsI18n, type TeamsI18n, type TeamsKey } from '../i18n/translator';

const LEVELS: readonly (readonly [NotificationLevel, TeamsKey, TeamsKey])[] = [
	['all', 'settings.levelAll', 'settings.levelAllText'],
	['mentions', 'settings.levelMentions', 'settings.levelMentionsText'],
	['off', 'settings.levelOff', 'settings.levelOffText'],
];

export type DesktopPermission = NotificationPermission | 'unsupported';

const PERMISSION_KEY = {
	granted: 'settings.permissionGranted',
	default: 'settings.permissionDefault',
	denied: 'settings.permissionDenied',
	unsupported: 'settings.permissionUnsupported',
} as const satisfies Record<DesktopPermission, TeamsKey>;

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
	i18n: TeamsI18n = teamsI18n(),
) {
	const { t } = i18n;
	return html`<fieldset class="notification-level">
			<legend>${t('settings.notifyAbout')}</legend>
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
						<span><strong>${t(label)}</strong><small>${t(description)}</small></span>
					</label>`,
			)}
		</fieldset>
		<h3>${t('settings.sound')}</h3>
		<label class="preference"
			><span>${t('settings.soundText')}</span
			><office-ui-switch
				aria-label=${t('settings.soundLabel')}
				.checked=${settings.sound}
				@change=${(event: Event) => handlers.change({ sound: (event.target as OfficeUiSwitch).checked })}
			></office-ui-switch
		></label>
		<h3>${t('settings.desktop')}</h3>
		<p class="description" role="status">${t(PERMISSION_KEY[permission])}</p>
		${
			permission === 'default'
				? html`<button type="button" class="primary" @click=${handlers.requestPermission}>
						${t('settings.enableDesktop')}
					</button>`
				: ''
		}
		<p class="description">${t('settings.dndNote')}</p>`;
}
