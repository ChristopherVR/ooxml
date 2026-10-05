// Two ways to use the React binding on one page:
//   left   <Teams />      the complete UI (the <teams-app> element)
//   right  useTeams()     your own markup over the raw client: state in, actions out
import { Teams, useTeams, type TeamsClientOptions } from 'openteams-react-viewer';
import { useMemo, useState } from 'react';

const params = new URLSearchParams(location.search);
const name = params.get('name') ?? 'Ada';
const userId = params.get('id') ?? `demo-${name.toLowerCase()}`;
const workspaceId = params.get('room') ?? 'react-demo';
const config = {
	mode: 'server' as const,
	syncUrl: 'ws://127.0.0.1:8787/sync',
	signalingUrl: 'ws://127.0.0.1:8787/signal',
	iceServers: [{ urls: 'stun:stun.l.google.com:19302' }],
};

export function App() {
	return (
		<div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) 22em', height: '100%' }}>
			<Teams workspaceId={workspaceId} userName={name} userId={userId} config={config} />
			<CustomPanel />
		</div>
	);
}

/** A tiny hand-rolled UI: no <teams-app>, only the hook. */
function CustomPanel() {
	const options = useMemo<TeamsClientOptions>(
		() => ({ workspaceId, user: { id: `${userId}-panel`, name: `${name} (panel)` }, config }),
		[],
	);
	const { client, state } = useTeams(options);
	const [text, setText] = useState('');
	if (!state || !client) return <aside style={{ padding: 16 }}>Connecting…</aside>;
	return (
		<aside style={{ padding: 16, borderLeft: '1px solid #ddd', overflow: 'auto' }}>
			<h3 style={{ marginTop: 0 }}>useTeams() panel</h3>
			<p style={{ color: '#666' }}>
				{state.status}, {state.people.length} online
			</p>
			<ul style={{ paddingLeft: 18 }}>
				{state.channels.map((c) => (
					<li key={c.id}>
						<button
							onClick={() => client.select(c.id)}
							style={{ fontWeight: c.id === state.selectedChannelId ? 700 : 400 }}
						>
							# {c.name} {c.unread ? `(${c.unread})` : ''}
						</button>
					</li>
				))}
			</ul>
			<ol style={{ paddingLeft: 18 }}>
				{state.messages.slice(-6).map((m) => (
					<li key={m.id}>
						<b>{m.authorName}:</b> {m.deleted ? <i>deleted</i> : m.text}
					</li>
				))}
			</ol>
			<form
				onSubmit={(e) => {
					e.preventDefault();
					void client.send({ text });
					setText('');
				}}
			>
				<input value={text} onChange={(e) => setText(e.target.value)} placeholder="Say something" />
			</form>
		</aside>
	);
}
