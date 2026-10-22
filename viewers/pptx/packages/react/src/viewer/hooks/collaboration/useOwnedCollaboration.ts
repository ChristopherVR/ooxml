import { needsOwnedSession, openOwnedCollaboration } from 'ooxml-ui/pptx';
import { useEffect, useMemo, useState } from 'react';

import type { CollaborationConfig, ConnectionStatus } from './types';

export interface OwnedCollaboration {
	/** The config to run: the input, or with the owned session as its `externalSession`. */
	config: CollaborationConfig | undefined;
	/** The session is being opened (or failed): nothing may connect with the input yet. */
	holding: boolean;
	/** `connecting` while opening, `error` when it could not open; otherwise `undefined`. */
	status: Extract<ConnectionStatus, 'connecting' | 'error'> | undefined;
}

type Opened = { source: CollaborationConfig; prepared: CollaborationConfig } | { failed: true };

/**
 * Opens the session a `provider` or same-browser (BroadcastChannel) config needs and destroys it
 * when the config changes, collaboration stops or the viewer unmounts. Legacy websocket and
 * WebRTC configs pass through untouched. Keyed on the connection and identity fields so a host
 * re-creating an equal config each render keeps one session.
 */
export function useOwnedCollaboration(config: CollaborationConfig | undefined): OwnedCollaboration {
	const needs = Boolean(config && needsOwnedSession(config));
	const [opened, setOpened] = useState<Opened | null>(null);
	useEffect(() => {
		if (!config || !needsOwnedSession(config)) {
			return;
		}
		const owned = openOwnedCollaboration(config);
		let live = true;
		owned.ready.then(
			(prepared) => {
				if (live && prepared) {
					setOpened({ source: config, prepared });
				}
			},
			() => {
				if (live) {
					setOpened({ failed: true });
				}
			},
		);
		return () => {
			live = false;
			owned.release();
			setOpened(null);
		};
		// Transport and identity only; see the hook comment.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [
		needs,
		config?.roomId,
		config?.provider,
		config?.serverUrl,
		config?.transport,
		config?.role,
		config?.userName,
		config?.userColor,
		config?.userAvatar,
	]);
	return useMemo((): OwnedCollaboration => {
		if (!needs) {
			return { config, holding: false, status: undefined };
		}
		if (opened && 'prepared' in opened) {
			return {
				config: { ...config!, externalSession: opened.prepared.externalSession! },
				holding: false,
				status: undefined,
			};
		}
		return { config, holding: true, status: opened ? 'error' : 'connecting' };
	}, [needs, config, opened]);
}
