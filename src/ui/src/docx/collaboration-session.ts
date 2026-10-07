import { emit } from './events';
import type { EditorView } from 'prosemirror-view';
import { CollaborationClient, type CollaborationConfig, type StepBatch } from 'ooxml-core/docx/ui';
import { createCollaborationIdGenerator } from 'ooxml-core/docx/ui';
import { EditorPresence } from './editor-presence';
import { WordYjsCollaboration, type WordYjsOptions } from 'ooxml-core/docx/ui';
import type { CollabSession } from 'ooxml-core/collab';
import type { Node } from 'prosemirror-model';

/** One editor's collaboration state: step client, presence and collaborative id generator. */
export class CollaborationSession {
	client?: CollaborationClient | undefined;
	yjs?: WordYjsCollaboration | undefined;
	presence?: EditorPresence | undefined;
	ids?: ((kind: string) => string) | undefined;
	private sendScheduled = false;

	constructor(
		private readonly element: HTMLElement,
		private readonly getView: () => EditorView | undefined,
	) {}

	get active(): boolean {
		return Boolean(this.client || this.yjs);
	}

	start(config: CollaborationConfig): void {
		if (this.active) throw new Error('Stop the current collaboration session first.');
		this.client = new CollaborationClient(config);
		this.presence = new EditorPresence(config, this.element, this.getView);
		this.ids = createCollaborationIdGenerator(config.clientId);
	}

	startYjs(session: CollabSession, doc: Node, options: WordYjsOptions): void {
		if (this.active) throw new Error('Stop the current collaboration session first.');
		this.yjs = new WordYjsCollaboration(session, doc, options);
		session.awareness.setLocalStateField('user', {
			name: session.identity.userName,
			color: session.identity.userColor,
		});
		this.ids = createCollaborationIdGenerator(String(session.clientId));
	}

	stop(): void {
		this.yjs?.destroy();
		this.yjs = undefined;
		this.client = undefined;
		this.presence?.leave();
		this.presence = undefined;
		this.ids = undefined;
	}

	/** Coalesces pending steps into one `collaboration-send` event per microtask. */
	scheduleSend(pending: () => StepBatch | null): void {
		if (!this.client || this.sendScheduled) return;
		this.sendScheduled = true;
		queueMicrotask(() => {
			this.sendScheduled = false;
			if (!this.element.isConnected) return;
			const batch = pending();
			if (batch) emit(this.element, 'collaboration-send', batch);
		});
	}
}
