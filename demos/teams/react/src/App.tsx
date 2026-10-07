// Two ways to use the React binding on one page:
//   left   <Teams />      the complete UI (the <teams-app> element)
//   right  useTeams()     your own markup over the raw client: state in, actions out
import { Teams, useTeams, type TeamsClientOptions } from 'openteams-react-viewer';
import type { TeamsServerConfig } from 'ooxml-core/teams';
import { useMemo, useState } from 'react';

const params = new URLSearchParams(location.search);
const name = params.get('name') ?? 'Ada';
const userId = params.get('id') ?? `demo-${name.toLowerCase()}`;
const workspaceId = params.get('room') ?? 'react-demo';
const iceServers = [{ urls: 'stun:stun.l.google.com:19302' }];
// The GitHub Pages build (VITE_TEAMS_STATIC=1) has no server behind it: it runs in local mode, where
// tabs of this browser (and the two clients on this page) share state over BroadcastChannel.
const staticSite = import.meta.env.VITE_TEAMS_STATIC === '1';
const config: TeamsServerConfig =
	staticSite || params.get('local')
		? { mode: 'local', iceServers }
		: {
				mode: 'server',
				syncUrl: 'ws://127.0.0.1:8787/sync',
				signalingUrl: 'ws://127.0.0.1:8787/signal',
				iceServers,
			};

export function App() {
	// A file preview does not mount the raw-hook client or the demo notice.
	if (params.get('openteams-file') === '1')
		return <Teams workspaceId={workspaceId} userName={name} userId={userId} config={config} />;
	return (
		<div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
			{staticSite ? <StaticNotice /> : null}
			<div
				style={{
					display: 'grid',
					gridTemplateColumns: 'minmax(0, 1fr) 22em',
					flex: 1,
					minHeight: 0,
				}}
			>
				<Teams workspaceId={workspaceId} userName={name} userId={userId} config={config} />
				<CustomPanel />
			</div>
		</div>
	);
}

/** Shown on the Pages build only: what runs where, and how to try two people. */
function StaticNotice() {
	const [open, setOpen] = useState(true);
	if (!open) return null;
	const other = name.toLowerCase() === 'bob' ? 'Ada' : 'Bob';
	const next = new URLSearchParams(params);
	next.set('name', other);
	next.delete('id');
	return (
		<div className="static-notice" role="note">
			<p>
				<strong>This demo runs entirely in your browser.</strong> There is no server behind this
				page: the full UI on the left and the <code>useTeams()</code> panel on the right are two
				clients that sync over BroadcastChannel, and nothing leaves this browser. Real use needs a
				server:{' '}
				<a href="/ooxml/teams/server" target="_top">
					run your own
				</a>
				.
			</p>
			<a href={`?${next.toString()}`} target="_blank" rel="noopener">
				Open a second tab as {other}
			</a>
			<button type="button" onClick={() => setOpen(false)}>
				Dismiss
			</button>
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
