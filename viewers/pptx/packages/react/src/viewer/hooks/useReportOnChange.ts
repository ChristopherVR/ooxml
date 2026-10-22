import { useEffect, useRef } from 'react';

/**
 * Calls `report(value)` each time `value` changes after the first render, never for the value
 * the component mounted with. Event callbacks that describe a transition (autosave status,
 * presentation mode) use it so a host hears about changes, not about the initial state.
 */
export function useReportOnChange<T>(
	value: T,
	report: ((value: T) => void) | undefined,
	initial: T = value,
): void {
	const last = useRef(initial);
	const callback = useRef(report);
	useEffect(() => {
		callback.current = report;
	}, [report]);
	useEffect(() => {
		if (Object.is(last.current, value)) {
			return;
		}
		last.current = value;
		callback.current?.(value);
	}, [value]);
}
