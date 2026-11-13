import { useEffect, useRef } from 'react';

/**
 * A toolbar control or canvas gesture opened an internal dialog by flag, which the building
 * blocks never render: when the flag turns on, `handle` clears it and raises the host request.
 *
 * The effect runs on the flag alone. Raising a request notifies the session's store, which
 * `useSyncExternalStore` re-renders synchronously; React 18 does that render before the flag's own
 * reset lands, so an effect that also depended on the render's objects (`dialogs`, `session`) ran
 * again with the flag still on, raised again and looped ("Maximum update depth exceeded").
 */
export function useDialogFlagRequest(isOpen: boolean, handle: () => void): void {
	const handleRef = useRef(handle);
	handleRef.current = handle;
	useEffect(() => {
		if (isOpen) {
			handleRef.current();
		}
	}, [isOpen]);
}
