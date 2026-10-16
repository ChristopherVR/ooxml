import { html } from 'lit';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';
import type { TeamsTheme } from './teams-settings';

const THEME_KEY = {
	system: 'settings.themeSystem',
	light: 'settings.themeLight',
	dark: 'settings.themeDark',
} as const satisfies Record<TeamsTheme, string>;

/** Decorative workspace previews accompany the existing theme buttons. */
export function themeChoices(
	current: TeamsTheme,
	change: (theme: TeamsTheme) => void,
	i18n: TeamsI18n = teamsI18n(),
) {
	return html`<div class="theme-options" role="group" aria-label=${i18n.t('settings.theme')}>
		${(['system', 'light', 'dark'] as const).map(
			(theme) =>
				html`<button
					type="button"
					aria-pressed=${String(current === theme)}
					@click=${() => change(theme)}
				>
					<span class="theme-swatch" data-theme=${theme} aria-hidden="true">
						<span class="preview-toolbar"></span>
						<span class="preview-rail"></span>
						<span class="preview-sidebar"><i></i><i></i><i></i></span>
						<span class="preview-conversation"><i></i><i></i><i></i></span>
					</span>
					<span>${i18n.t(THEME_KEY[theme])}</span>
				</button>`,
		)}
	</div>`;
}
