// @vitest-environment happy-dom
/**
 * A peer-to-peer session that cannot start says why: the real collaboration
 * provider runs the shared session controller against a webrtc config that
 * lacks its invitation secret, and the status pill and the Share dialog show
 * the translated reason instead of only "Connection error".
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import type { CollaborationConfig } from '../../hooks/collaboration/types';
import { ShareDialog } from '../ShareDialog';
import { CollaborationProvider, useCollaboration } from './CollaborationProvider';
import { CollaborationStatusIndicator } from './CollaborationStatusIndicator';

const MISSING_SECRET = translationsEn['pptx.collaboration.failure.missingSecret'];

function StatusPill(): React.ReactElement | null {
	const collab = useCollaboration();
	return collab ? (
		<CollaborationStatusIndicator
			status={collab.status}
			connectedCount={collab.connectedCount}
			failure={collab.failure}
		/>
	) : null;
}

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

async function waitFor(check: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 100 && !check(); attempt += 1) {
		await act(
			() =>
				new Promise<void>((resolve) => {
					setTimeout(resolve, 10);
				}),
		);
	}
}

describe('collaboration start failure (React)', () => {
	it('shows why a peer-to-peer session without its secret could not start', async () => {
		const config: CollaborationConfig = {
			roomId: 'room-1',
			serverUrl: '',
			userName: 'Ada',
			transport: 'webrtc',
		};
		act(() => {
			root.render(
				<CollaborationProvider config={config} canvasWidth={960} canvasHeight={540}>
					<StatusPill />
					<ShareDialog open onClose={() => {}} />
				</CollaborationProvider>,
			);
		});
		const banner = (): Element | null =>
			document.querySelector('[data-testid="collaboration-failure"]');
		await waitFor(() => banner() !== null);
		expect(banner()?.textContent).toBe(MISSING_SECRET);
		const pill = container.querySelector('[data-testid="collaboration-status"]');
		expect(pill?.textContent).toContain(translationsEn['pptx.collaboration.status.error']);
		expect(pill?.getAttribute('title')).toBe(MISSING_SECRET);
	});

	it('says why a peer-to-peer invitation without its key cannot join', () => {
		act(() => {
			root.render(<ShareDialog open onClose={() => {}} defaultUserName='Ada' />);
		});
		const joinTab = Array.from(document.querySelectorAll<HTMLButtonElement>('[role="tab"]'))[1];
		act(() => joinTab?.click());
		const field = document.querySelector<HTMLInputElement>('#share-invitation');
		act(() => {
			// React tracks the value through the native setter; an input event then commits it.
			Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(
				field,
				'https://x.test/?room=deck-1&transport=webrtc',
			);
			field?.dispatchEvent(new Event('input', { bubbles: true }));
		});
		expect(document.querySelector('[data-testid="collaboration-failure"]')?.textContent).toBe(
			MISSING_SECRET,
		);
	});
});
