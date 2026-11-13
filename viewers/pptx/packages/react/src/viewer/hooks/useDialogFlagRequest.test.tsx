// @vitest-environment happy-dom
/**
 * A dialog flag that becomes a host request must be handled once, even though raising the request
 * re-renders synchronously through `useSyncExternalStore` before the flag's reset lands (React 18
 * renders that sync update first; an effect keyed on the render's objects then looped).
 */
import { createControllerStore } from 'ooxml-ui/pptx';
import React, { act, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, test, vi } from 'vitest';

import { useController } from './useController';
import { useDialogFlagRequest } from './useDialogFlagRequest';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;

let container: HTMLDivElement | undefined;
afterEach(() => {
	container?.remove();
});

test('handles the flag once while the raised request re-renders synchronously', async () => {
	const requests = createControllerStore<{ kind: string } | null>(null);
	const raised = vi.fn();
	function Harness() {
		const [open, setOpen] = useState(true);
		useController(requests);
		// A new object every render, as the viewer's `dialogs` result is.
		const dialogs = { close: () => setOpen(false) };
		useDialogFlagRequest(open, () => {
			dialogs.close();
			raised();
			requests.set({ kind: 'smartArt' });
		});
		return null;
	}
	container = document.createElement('div');
	document.body.append(container);
	const root = createRoot(container);
	await act(async () => root.render(<Harness />));
	expect(raised).toHaveBeenCalledOnce();
	expect(requests.getSnapshot()).toEqual({ kind: 'smartArt' });
	await act(async () => root.unmount());
});
