// @vitest-environment happy-dom
/**
 * `useRecentColorCommit`: the ref callback that pushes a native colour input's
 * value into the "Recent Colors" list on the native `change` event (a committed
 * pick) and never on `input` (a drag tick through the OS picker).
 */
import React, { StrictMode, act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { RecentColorsProvider } from './RecentColorsContext';
import { useRecentColorCommit } from './useRecentColorCommit';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function fire(input: HTMLInputElement, type: 'input' | 'change', value: string): void {
	act(() => {
		Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(input, value);
		input.dispatchEvent(new Event(type, { bubbles: true }));
	});
}

/** Two colour inputs sharing ONE callback, as a list of gradient stops does. */
function Pickers({ onApply }: { onApply: (hex: string) => void }): React.ReactElement {
	const commitRecentColor = useRecentColorCommit();
	return (
		<>
			{['#000000', '#ffffff'].map((initial) => (
				<input
					key={initial}
					type='color'
					defaultValue={initial}
					ref={commitRecentColor}
					onChange={(event) => onApply(event.target.value)}
				/>
			))}
		</>
	);
}

function renderPickers(pushColor: (hex: string) => void, onApply = vi.fn(), strict = false): void {
	const tree = (
		<RecentColorsProvider value={{ recentColors: [], pushColor }}>
			<Pickers onApply={onApply} />
		</RecentColorsProvider>
	);
	act(() => root.render(strict ? <StrictMode>{tree}</StrictMode> : tree));
}

function inputs(): HTMLInputElement[] {
	return Array.from(container.querySelectorAll<HTMLInputElement>('input[type="color"]'));
}

describe('useRecentColorCommit', () => {
	it('pushes nothing on input events and the settled value on change', () => {
		const pushColor = vi.fn();
		const onApply = vi.fn();
		renderPickers(pushColor, onApply);
		const [first] = inputs();

		fire(first, 'input', '#110000');
		fire(first, 'input', '#220000');
		expect(onApply).toHaveBeenCalledTimes(2);
		expect(pushColor).not.toHaveBeenCalled();

		fire(first, 'change', '#220000');
		expect(pushColor).toHaveBeenCalledExactlyOnceWith('#220000');
	});

	it('serves several inputs with one callback, pushing each one its own value', () => {
		const pushColor = vi.fn();
		renderPickers(pushColor);
		const [first, second] = inputs();

		fire(second, 'change', '#00ff00');
		fire(first, 'change', '#ff0000');
		expect(pushColor.mock.calls).toStrictEqual([['#00ff00'], ['#ff0000']]);
	});

	it('pushes once per change after re-renders and under StrictMode', () => {
		const pushColor = vi.fn();
		renderPickers(pushColor, vi.fn(), true);
		renderPickers(pushColor, vi.fn(), true);
		const [first] = inputs();

		fire(first, 'change', '#123456');
		expect(pushColor).toHaveBeenCalledExactlyOnceWith('#123456');
	});

	it('pushes through the latest pushColor after the provider value changes', () => {
		const stale = vi.fn();
		const latest = vi.fn();
		renderPickers(stale);
		renderPickers(latest);

		fire(inputs()[0], 'change', '#abcdef');
		expect(stale).not.toHaveBeenCalled();
		expect(latest).toHaveBeenCalledExactlyOnceWith('#abcdef');
	});
});
