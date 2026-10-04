import { defineTeamsApp, type TeamsApp } from 'teams-viewer';

defineTeamsApp();
const app = document.querySelector<TeamsApp>('teams-app')!;

// ?name=Ada&id=ada opens the demo as that person, so two tabs can talk to each other.
const params = new URLSearchParams(location.search);
const name = params.get('name');
if (name) {
	app.setAttribute('user-name', name);
	app.setAttribute('user-id', params.get('id') ?? `demo-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`);
}
if (params.get('room')) app.setAttribute('workspace-id', params.get('room')!);

// First visit: use the reference server from `bun run dev`. Settings (Server button) override it
// and are remembered; ?local=1 forces the serverless same-browser mode.
let hasSaved = false;
try {
	hasSaved = localStorage.getItem('teams:config') !== null;
} catch {
	// storage blocked: fall through to the default
}
if (params.get('local')) app.config = { mode: 'local', iceServers: [{ urls: 'stun:stun.l.google.com:19302' }] };
else if (!hasSaved)
	app.config = {
		mode: 'server',
		syncUrl: 'ws://127.0.0.1:8787/sync',
		signalingUrl: 'ws://127.0.0.1:8787/signal',
		iceServers: [{ urls: 'stun:stun.l.google.com:19302' }],
	};

app.addEventListener('teams-open-file', (event) => {
	const { attachment, url } = (event as CustomEvent).detail;
	console.info('open file', attachment.kind, url);
});
