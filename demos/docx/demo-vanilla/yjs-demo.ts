import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import type { DocxEditorElement } from 'docx-web-component';

/** The demo host owns provider lifecycle; every framework mounts the same Word editor. */
export function startYjsDemo(peers: DocxEditorElement[]): void {
	document.getElementById('collaboration-description')!.textContent =
		'Two editors share a document through local Yjs providers. Pause delivery to try concurrent edits, then resume to merge them. Nothing is sent over the network.';
	document.getElementById('collaboration-limitations')!.textContent =
		'Colored selections use relative cursor positions. The demo shares text, formatting and inserted pictures in memory. Production hosts provide transport, access control and storage. Editing comments, notes and other stories together remains unsupported.';
	let paused = false;
	const hub = createMemoryHub({ filter: () => !paused });
	const sessions = peers.map((_peer, index) =>
		createCollabSession({
			roomId: 'word-yjs-demo',
			provider: transportProvider({ transport: hub.createTransport('word-yjs-demo') }),
			user: { name: index === 0 ? 'Alice' : 'Bob' },
			heartbeatMs: 0,
		}),
	);
	peers.forEach((peer, index) =>
		peer.startYjsCollaboration(sessions[index]!, {
			documentId: 'shared-demo-document',
			initializeIfEmpty: index === 0,
		}),
	);
	peers.forEach((peer, index) => {
		const prefix = `peer-${index === 0 ? 'a' : 'b'}`;
		const name = document.getElementById(`${prefix}-name`) as HTMLInputElement;
		const color = document.getElementById(`${prefix}-color`) as HTMLSelectElement;
		const publish = () => peer.publishPresence({ name: name.value, color: color.value });
		name.addEventListener('input', publish);
		color.addEventListener('change', publish);
		publish();
	});
	const status = document.getElementById('collaboration-status')!;
	const renderStatus = () => {
		status.textContent = paused ? 'Delivery paused' : 'Synced · Yjs';
	};
	const pause = document.getElementById('pause-delivery')!;
	pause.addEventListener('click', () => {
		paused = !paused;
		pause.textContent = paused ? 'Resume delivery' : 'Pause delivery';
		pause.setAttribute('aria-pressed', String(paused));
		if (!paused) for (const peer of peers) peer.resyncCollaboration();
		renderStatus();
	});
	for (const [label, action] of [
		['Reconnect providers', () => peers.forEach((peer) => peer.reconnectCollaboration())],
		['Resync providers', () => peers.forEach((peer) => peer.resyncCollaboration())],
	] as const) {
		const button = document.createElement('button');
		button.type = 'button';
		button.textContent = label;
		button.addEventListener('click', action);
		pause.after(button);
	}
	document.getElementById('peer-readonly')!.addEventListener('change', (event) => {
		peers[1]!.readOnly = (event.target as HTMLInputElement).checked;
	});
	window.addEventListener(
		'pagehide',
		() => {
			for (const peer of peers) peer.stopCollaboration();
			for (const session of sessions) session.destroy();
		},
		{ once: true },
	);
	renderStatus();
}
