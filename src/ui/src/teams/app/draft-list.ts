import { html } from 'lit';
import type { DraftContext, TeamsClient, TeamsState } from 'ooxml-core/teams';

export function draftList(
	state: TeamsState,
	client: TeamsClient,
	open: (context: DraftContext) => void,
) {
	return html`<section class="followed-threads" aria-label="Drafts">
		<h1>Drafts</h1>
		<p>Unsent messages on this device. After reloading, draft attachments must be reattached.</p>
		<ul>
			${
				state.drafts.length
					? state.drafts.map((draft) => {
							const channel = state.channels.find((item) => item.id === draft.context.channelId)!;
							return html`<li>
								<button
									type="button"
									aria-label=${`Resume draft in ${channel.name}: ${draft.text.slice(0, 80) || 'Attachments'}`}
									@click=${() => open(draft.context)}
								>
									<strong># ${channel.name}</strong><span>${draft.text || 'Attachments'}</span
									><small
										>${draft.context.editId ? 'Edit' : draft.context.threadId ? 'Thread reply' : 'New post'}
										· ${draft.files.length + draft.missingFiles.length} attachments</small
									></button
								><button
									type="button"
									@click=${() => {
										if (globalThis.confirm?.('Discard this unsent draft?') !== false)
											client.discardDraft(draft.context);
									}}
								>
									Discard draft
								</button>
							</li>`;
						})
					: html`<li>No unsent drafts.</li>`
			}
		</ul>
	</section>`;
}
