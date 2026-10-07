import { html, nothing } from 'lit';

export function workbookActions(
	state: { editing: boolean; saving: boolean; canShare: boolean },
	actions: { toggle(): void; share(): void; download(): void },
) {
	return html`<button type="button" ?disabled=${state.saving} @click=${actions.toggle}>
			${state.editing ? 'View workbook' : 'Edit workbook'}
		</button>
		<button type="button" ?disabled=${state.saving} @click=${actions.download}>
			Download workbook copy
		</button>
		${state.canShare ? html`<button type="button" ?disabled=${state.saving || !state.editing} @click=${actions.share}>${state.saving ? 'Saving copy...' : 'Save copy to channel'}</button>` : nothing}`;
}
