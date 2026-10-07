import { html } from 'lit';
import type { FollowedThread, TeamsClient, TeamsState } from 'ooxml-core/teams';

export function followedThreads(
	state: TeamsState,
	client: TeamsClient,
	open: (thread: FollowedThread) => void,
) {
	return html`<section class="followed-threads" aria-label="Followed threads">
		<h1>Followed threads</h1>
		<p>Threads you follow on this device. Notifications are not enabled.</p>
		<ul>
			${
				state.followedThreads.length
					? state.followedThreads.map(
							(thread) => html`<li>
								<button
									type="button"
									@click=${() => open(thread)}
									aria-label=${`Open thread in ${thread.channelName}: ${thread.root.deleted ? 'Deleted message' : thread.root.text}`}
								>
									<strong># ${thread.channelName}</strong
									><span
										>${thread.root.deleted ? 'This message was deleted' : thread.root.text || 'Shared file'}</span
									>
									<small
										>${thread.replyCount} ${thread.replyCount === 1 ? 'reply' : 'replies'} ·
										${new Date(thread.updatedAt).toLocaleString()}</small
									></button
								><button
									type="button"
									@click=${() => client.followThread(thread.channelId, thread.root.id, false)}
								>
									Unfollow thread
								</button>
							</li>`,
						)
					: html`<li>No followed threads yet. Open a thread and choose Follow thread.</li>`
			}
		</ul>
	</section>`;
}
