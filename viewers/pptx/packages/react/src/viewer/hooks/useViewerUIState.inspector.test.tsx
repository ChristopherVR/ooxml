// @vitest-environment happy-dom
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it } from 'vitest';

import { useViewerUIState } from './useViewerUIState';

describe('initial inspector visibility (#61)', () => {
	it('starts closed on request, preserves toggles and ignores later initial-option changes', async () => {
		let state!: ReturnType<typeof useViewerUIState>;
		function Harness({ open }: { open: boolean }) {
			state = useViewerUIState(open);
			return null;
		}
		const root = createRoot(document.createElement('div'));
		try {
			await act(async () => root.render(<Harness open={false} />));
			expect(state.isInspectorPaneOpen).toBe(false);
			await act(async () => state.setIsInspectorPaneOpen(true));
			await act(async () => root.render(<Harness open={false} />));
			expect(state.isInspectorPaneOpen).toBe(true);
		} finally {
			await act(async () => root.unmount());
		}
	});
});
