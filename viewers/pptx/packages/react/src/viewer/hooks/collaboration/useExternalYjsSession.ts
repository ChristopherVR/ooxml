import {
	borrowExternalCollaborationAwareness,
	observeExternalCollaborationSession,
	registerCollaborationTeardown,
} from 'ooxml-ui/pptx';
import type { BorrowedCollaborationAwareness, ExternalCollaborationSession } from 'ooxml-ui/pptx';
import { useEffect, useRef, useState } from 'react';

import type { UseYjsProviderResult } from './useYjsProvider';

type ExternalYjsSessionResult = Omit<UseYjsProviderResult, 'awareness'> & {
	awareness: BorrowedCollaborationAwareness | null;
};

const inactive: ExternalYjsSessionResult = {
	status: 'disconnected',
	awareness: null,
	doc: null,
	clientId: null,
	synced: false,
	// An external transport can only be retried, or refuse an update, by its owner.
	retry: () => {},
	refusal: null,
};

/** Attach to host resources without owning their network or lifetime. */
export function useExternalYjsSession(
	session: ExternalCollaborationSession | undefined,
	leaveMessageOrigin?: string | readonly string[],
): ExternalYjsSessionResult {
	const origin = useRef(leaveMessageOrigin);
	origin.current = leaveMessageOrigin;
	const [binding, setBinding] = useState<{
		session: ExternalCollaborationSession;
		value: ExternalYjsSessionResult;
	} | null>(null);

	useEffect(() => {
		if (!session) {
			return;
		}
		let detach: (() => void) | undefined;
		const leave = (): void => {
			detach?.();
			detach = undefined;
			setBinding(null);
		};
		const attach = (): void => {
			leave();
			const borrowed = borrowExternalCollaborationAwareness(session.awareness);
			try {
				const unsubscribe = observeExternalCollaborationSession(session, (snapshot) => {
					setBinding({
						session,
						value: {
							...snapshot,
							doc: session.doc,
							awareness: borrowed.awareness,
							clientId: session.awareness.clientID,
							retry: inactive.retry,
							refusal: null,
						},
					});
				});
				detach = () => {
					unsubscribe();
					borrowed.dispose();
				};
			} catch {
				borrowed.dispose();
				setBinding({ session, value: { ...inactive, status: 'error' } });
			}
		};
		attach();
		const unregister = registerCollaborationTeardown({
			leave,
			rejoin: attach,
			leaveOnBeforeUnload: false,
			leaveMessageOrigin: () => origin.current,
		});
		return () => {
			unregister();
			leave();
		};
	}, [session]);

	return binding && binding.session === session ? binding.value : inactive;
}
