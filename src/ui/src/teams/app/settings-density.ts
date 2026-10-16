import { html } from 'lit';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';

export type ChatDensity = 'comfy' | 'compact';

/** Native radios retain keyboard and screen-reader behavior. */
export function densityChoices(
	density: ChatDensity,
	change: (density: ChatDensity) => void,
	i18n: TeamsI18n = teamsI18n(),
) {
	const { t } = i18n;
	return html`<fieldset class="density-options">
		<legend>${t('settings.density')}</legend>
		${(['comfy', 'compact'] as const).map(
			(value) =>
				html`<label class="density-choice">
					<input
						type="radio"
						name="chat-density"
						value=${value}
						.checked=${density === value}
						@change=${() => change(value)}
					/>
					<span class="density-preview" data-density=${value} aria-hidden="true">
						${[0, 1, 2].map(
							() =>
								html`<span class="preview-message"
									><i></i><span><b></b><b></b></span
								></span>`,
						)}
					</span>
					<span class="density-caption"
						><strong>${t(value === 'comfy' ? 'settings.comfy' : 'settings.compact')}</strong>
						<small
							>${t(value === 'comfy' ? 'settings.comfyText' : 'settings.compactText')}</small
						></span
					>
				</label>`,
		)}
	</fieldset>`;
}
