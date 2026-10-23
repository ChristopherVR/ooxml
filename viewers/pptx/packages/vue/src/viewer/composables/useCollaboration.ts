/**
 * Vue session lifecycle for owned providers or a borrowed host connection.
 * The room lifecycle is the shared `createCollaborationSessionController`;
 * document readiness is shared across bindings; Vue watches and presence have
 * focused composables.
 */
import type { CollaborationConfig, CollaborationRole, ConnectionStatus } from 'ooxml-ui/pptx';
import {
	buildCollaborationState,
	createCollaborationSessionController,
	resolveCollaborationConfig,
	resolveCollaborationShellState,
	writeBackReport,
} from 'ooxml-ui/pptx';
import { computed, onScopeDispose, toValue, watch } from 'vue';
import type { Ref } from 'vue';

import { collabModules } from './collaboration-modules';
import type { UseCollaborationOptions, UseCollaborationResult } from './collaboration-types';
import { useCollaborationDocumentSync } from './useCollaborationDocumentSync';
import { useCollaborationPresence } from './useCollaborationPresence';
import { useController } from './useController';

export type {
	RemotePresence,
	UseCollaborationOptions,
	UseCollaborationResult,
} from './collaboration-types';

export function useCollaboration(options: UseCollaborationOptions): UseCollaborationResult {
	const hostConfig = (): CollaborationConfig | undefined =>
		resolveCollaborationConfig(toValue(options.collaboration));
	// The document layer reports a borrowed session's readiness through this.
	const statusPort = {
		get value(): ConnectionStatus {
			return session.getSnapshot().status;
		},
		set value(next: ConnectionStatus) {
			session.setStatus(next);
		},
	};
	const presence = useCollaborationPresence(options);
	const document = useCollaborationDocumentSync(options, statusPort);
	const session = createCollaborationSessionController({
		modules: collabModules,
		gate: document.gate,
		onBegin: (config) => document.begin(config),
		onOpen: (live, config) => {
			presence.start(live.awareness, config);
			// Observe and catch up before the first-write gate opens: a provider
			// can already contain the room while its async factory is resolving.
			document.attach(live.doc, live.factories);
		},
		onClose: () => {
			document.stop();
			presence.stop();
		},
	});
	const sessionState = useController(session);
	const status = computed(() => sessionState.value.status) as Ref<ConnectionStatus>;
	const connected = computed(() => status.value === 'connected');
	const active = computed(() => sessionState.value.active) as Ref<boolean>;
	// The host's config stays the identity of the running session (the watcher compares it).
	const activeCollaboration = computed(
		() => sessionState.value.config,
	) as Ref<CollaborationConfig | null>;
	const activeRole = computed(() => sessionState.value.config?.role) as Ref<
		CollaborationRole | undefined
	>;
	const shellState = computed(() =>
		resolveCollaborationShellState({
			authorizedCanEdit: toValue(options.canEdit) ?? true,
			configured: Boolean(hostConfig() || activeCollaboration.value),
			readOnly: document.readOnly.value,
			sourcePending: toValue(options.sourcePending) ?? false,
			sourceError: toValue(options.sourceError) ?? false,
			status: status.value,
			remoteUsers: presence.remoteUsers.value,
		}),
	);
	const state = computed(() =>
		buildCollaborationState({
			active: active.value,
			roomId: activeCollaboration.value?.roomId,
			status: status.value,
			synced: document.synced.value,
			clientId: presence.localClientId.value,
			userName: activeCollaboration.value?.userName,
			userColor: options.userColor ?? activeCollaboration.value?.userColor,
			remoteUsers: presence.remoteUsers.value,
			writeBack: writeBackReport(activeCollaboration.value, document.writeBackState.value),
		}),
	);
	const connectedCount = computed(
		() => presence.remotePresences.value.length + (active.value ? 1 : 0),
	);
	const start = (config: CollaborationConfig): Promise<void> => session.start(config);
	const stop = (): void => session.stop();
	const retry = (): Promise<void> => session.retry();
	watch(
		hostConfig,
		(config) => {
			if (config && config !== activeCollaboration.value) {
				void start(config);
			} else if (!config && activeCollaboration.value) {
				stop();
			}
		},
		{ immediate: true },
	);
	onScopeDispose(() => {
		session.destroy();
		document.dispose();
	});
	return {
		state,
		shellState,
		activeCollaboration,
		status,
		connected,
		active,
		activeRole,
		connectedCount,
		readOnly: document.readOnly,
		cursors: presence.cursors,
		remotePresences: presence.remotePresences,
		followedClientId: presence.followedClientId,
		followedSlideIndex: presence.followedSlideIndex,
		broadcasterSlideIndex: presence.broadcasterSlideIndex,
		setCursor: presence.setCursor,
		setSelection: presence.setSelection,
		setActiveSlide: presence.setActiveSlide,
		followUser: presence.followUser,
		start,
		stop,
		retry,
		livePatcher: document.livePatcher,
	};
}
