import {
	createCollaborationAuthority,
	PRESENCE_PALETTE,
	type DocxEditorElement,
	type PresenceMessage,
	type StepBatch,
} from 'docx-web-component';
import type { DocumentModel } from 'docx-core';

/**
 * A shared session between windows of one browser, whatever framework each demo was built with.
 *
 * The session is only a name: every window opens a BroadcastChannel called after it, so the React
 * demo can join a session started in the Vue demo. This is demo plumbing around the editor's
 * transport-neutral collaboration API, not a product feature: one window (the host, the one that
 * opened the sample) runs the in-memory reference authority and orders everybody's edits, and a
 * guest starts from the host's starting document and replays the accepted batches. The session
 * ends with its host, and it cannot reach another device. A real application brings its own
 * authority, transport and storage.
 */
type Message =
	| { type: 'join'; clientId: string }
	| { type: 'welcome'; to: string; model: DocumentModel; history: StepBatch[] }
	| { type: 'batch'; batch: StepBatch }
	| { type: 'accepted'; batch: StepBatch }
	| { type: 'presence'; message: PresenceMessage };

export interface SessionOptions {
	room: string;
	host: boolean;
	name: string;
	onStatus(text: string): void;
}

export function runSession(element: DocxEditorElement, options: SessionOptions): () => void {
	const { room, host, onStatus } = options;
	const channel = new BroadcastChannel(`docx-viewer-demo-${room}`);
	const clientId = `${host ? 'host' : 'guest'}-${Math.random().toString(36).slice(2, 8)}`;
	const send = (message: Message) => channel.postMessage(message);
	const profile = {
		name: options.name,
		color: PRESENCE_PALETTE[host ? 0 : 1 + Math.floor(Math.random() * 6)] ?? '#2563eb',
	};
	const waitingPresence = new Map<string, PresenceMessage>();
	let started = false;

	function retryPresence() {
		for (const [id, message] of waitingPresence) {
			if (element.receivePresence(message) !== 'out-of-order') waitingPresence.delete(id);
		}
	}
	function applyAccepted(batch: StepBatch) {
		const status = element.receiveCollaboration(batch);
		if (status !== 'applied' && status !== 'duplicate')
			onStatus(`Could not apply an edit: ${status}`);
		retryPresence();
	}
	element.addEventListener('presence-send', (event) => {
		send({ type: 'presence', message: (event as CustomEvent<PresenceMessage>).detail });
	});

	let authority: ReturnType<typeof createCollaborationAuthority> | undefined;
	let initial: DocumentModel | undefined;
	const history: StepBatch[] = [];

	function submit(batch: StepBatch) {
		if (!authority) return;
		const result = authority.submit(batch);
		if (result.status === 'rejected') {
			onStatus(`The authority rejected an edit: ${result.reason}`);
			return;
		}
		if (result.status === 'accepted') history.push(result.batch);
		applyAccepted(result.batch);
		send({ type: 'accepted', batch: result.batch });
	}

	function begin() {
		element.startCollaboration({ sessionId: room, clientId });
		element.publishPresence(profile);
		started = true;
	}

	element.addEventListener('collaboration-send', (event) => {
		const batch = (event as CustomEvent<StepBatch>).detail;
		if (host) submit(batch);
		else send({ type: 'batch', batch });
	});

	const early: StepBatch[] = [];
	channel.onmessage = (event: MessageEvent<Message>) => {
		const message = event.data;
		if (message.type === 'presence') {
			if (!started) return;
			if (element.receivePresence(message.message) === 'out-of-order')
				waitingPresence.set(message.message.clientId, message.message);
			return;
		}
		if (host) {
			if (message.type === 'join' && initial)
				send({ type: 'welcome', to: message.clientId, model: initial, history });
			if (message.type === 'batch') submit(message.batch);
			return;
		}
		if (message.type === 'welcome' && message.to === clientId && !started) {
			element.documentModel = message.model;
			begin();
			for (const batch of [...message.history, ...early]) applyAccepted(batch);
			early.length = 0;
			onStatus(`Joined session ${room}`);
		}
		if (message.type === 'accepted') {
			if (started) applyAccepted(message.batch);
			else early.push(message.batch);
		}
	};

	if (host) {
		initial = structuredClone(element.documentModel) as DocumentModel;
		authority = createCollaborationAuthority(initial, { sessionId: room });
		begin();
		onStatus(`Hosting session ${room}`);
	} else {
		onStatus(`Waiting for the host of session ${room}`);
		const ask = () => {
			if (!started) send({ type: 'join', clientId });
		};
		ask();
		const timer = setInterval(() => (started ? clearInterval(timer) : ask()), 1000);
	}
	return () => channel.close();
}
