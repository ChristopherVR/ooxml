// Settings shared by the Vue, Angular, Svelte and Solid demos: the same query parameters as the
// vanilla and React demos (?name=Ada&room=acme&local=1) and the same static-site notice.
import './theme-sync';
import './shared.css';
import type { TeamsServerConfig } from 'ooxml-core/teams';

const params = new URLSearchParams(location.search);
export const userName = params.get('name') ?? 'Ada';
export const userId =
	params.get('id') ?? `demo-${userName.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`;
export const workspaceId = params.get('room') ?? 'demo';
const iceServers = [{ urls: 'stun:stun.l.google.com:19302' }];
// The GitHub Pages build (VITE_TEAMS_STATIC=1) has no server behind it: local mode, where tabs and
// frames of this browser share chat, presence and calls over BroadcastChannel.
export const staticSite = import.meta.env.VITE_TEAMS_STATIC === '1';
export const config: TeamsServerConfig =
	staticSite || params.get('local')
		? { mode: 'local', iceServers }
		: {
				mode: 'server',
				syncUrl: 'ws://127.0.0.1:8787/sync',
				signalingUrl: 'ws://127.0.0.1:8787/signal',
				iceServers,
			};

/** Pages build only: say what runs where and offer a second person in a new tab. */
export function showStaticNotice(): void {
	if (!staticSite) return;
	const other = userName.toLowerCase() === 'bob' ? 'Ada' : 'Bob';
	const next = new URLSearchParams(params);
	next.set('name', other);
	next.delete('id');
	const notice = document.createElement('div');
	notice.className = 'static-notice';
	notice.setAttribute('role', 'note');
	const text = document.createElement('p');
	const strong = document.createElement('strong');
	strong.textContent = 'This demo runs entirely in your browser.';
	const link = document.createElement('a');
	link.href = '/teams-viewer/server';
	link.target = '_top';
	link.textContent = 'run your own';
	text.append(
		strong,
		' There is no server behind this page: chat, presence and calls travel only between tabs of this browser (BroadcastChannel). Real use needs a server: ',
		link,
		'.',
	);
	const second = document.createElement('a');
	second.href = `?${next.toString()}`;
	second.target = '_blank';
	second.rel = 'noopener';
	second.textContent = `Open a second tab as ${other}`;
	const dismiss = document.createElement('button');
	dismiss.type = 'button';
	dismiss.textContent = 'Dismiss';
	dismiss.addEventListener('click', () => notice.remove());
	notice.append(text, second, dismiss);
	document.body.prepend(notice);
}
