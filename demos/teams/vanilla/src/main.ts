import '../../theme-sync';
import { mountTeams, type TeamsElementProps } from 'openteams-vanilla-viewer';
import { currentHostClass, onHostClass, recordOpenFile } from '../../test-hooks';

// ?name=Ada&id=ada opens the demo as that person, so two tabs can talk to each other.
const params = new URLSearchParams(location.search);
const name = params.get('name');
const props: TeamsElementProps = {
	className: currentHostClass(),
	onOpenFile: (detail, event) => {
		recordOpenFile(detail, event);
		console.info('open file', detail.attachment.kind, detail.url);
	},
};
if (name) {
	props.userName = name;
	props.userId = params.get('id') ?? `demo-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`;
}
if (params.get('room')) props.workspaceId = params.get('room')!;

// The GitHub Pages build (VITE_TEAMS_STATIC=1) has no server behind it, so it starts in local mode:
// tabs of this browser share chat, presence and calls over BroadcastChannel and nothing leaves the
// browser. Settings the visitor saved with the Server button still win.
const staticSite = import.meta.env.VITE_TEAMS_STATIC === '1';
const iceServers = [{ urls: 'stun:stun.l.google.com:19302' }];

// First visit: use the reference server from `bun run dev`. Settings (Server button) override it
// and are remembered; ?local=1 forces the serverless same-browser mode.
let hasSaved = false;
try {
	hasSaved = localStorage.getItem('teams:config') !== null;
} catch {
	// storage blocked: fall through to the default
}
if (params.get('local') || (staticSite && !hasSaved)) props.config = { mode: 'local', iceServers };
else if (!hasSaved)
	props.config = {
		mode: 'server',
		syncUrl: 'ws://127.0.0.1:8787/sync',
		signalingUrl: 'ws://127.0.0.1:8787/signal',
		iceServers,
	};

// The whole app is one <teams-app>, mounted and kept in sync by the vanilla binding.
const teams = mountTeams(document.body, props);
onHostClass((className) => teams.update({ ...props, className }));

if (staticSite && params.get('openteams-file') !== '1') showStaticNotice();

/** Say that this page runs without a server, and offer a second person in a new tab. */
function showStaticNotice(): void {
	const notice = document.getElementById('static-notice');
	if (!notice) return;
	const other = name?.toLowerCase() === 'bob' ? 'Ada' : 'Bob';
	const link = notice.querySelector<HTMLAnchorElement>('[data-second-tab]');
	if (link) {
		const next = new URLSearchParams(params);
		next.set('name', other);
		next.delete('id');
		link.href = `?${next.toString()}`;
		link.textContent = `Open a second tab as ${other}`;
	}
	notice.querySelector('[data-dismiss]')?.addEventListener('click', () => notice.remove());
	notice.hidden = false;
}
