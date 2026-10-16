import { html, nothing } from 'lit';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';

export function workbookActions(
	state: { editing: boolean; saving: boolean; canShare: boolean },
	actions: { toggle(): void; share(): void; download(): void; cancel(): void },
	i18n: TeamsI18n = teamsI18n(),
) {
	const { t } = i18n;
	return html`<button type="button" ?disabled=${state.saving} @click=${actions.toggle}>
			${t(state.editing ? 'workbook.view' : 'workbook.edit')}
		</button>
		<button type="button" ?disabled=${state.saving} @click=${actions.download}>
			${t('workbook.download')}
		</button>
		${state.canShare ? html`<button type="button" ?disabled=${state.saving || !state.editing} @click=${actions.share}>${t(state.saving ? 'workbook.saving' : 'workbook.saveToChannel')}</button>` : nothing}
		${state.saving ? html`<button type="button" @click=${actions.cancel}>${t('workbook.cancelSave')}</button>` : nothing}`;
}
