// @vitest-environment happy-dom
/**
 * The host events every binding reports the same way (issue 151, item 5): the pieces of the
 * React wiring that do not need a mounted viewer.
 */
import type { PptxCoreProperties } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { usePropertyHandlers } from './usePropertyHandlers';
import type { UsePropertyHandlersInput } from './usePropertyHandlers';
import { useReportOnChange } from './useReportOnChange';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('useReportOnChange', () => {
	function Probe(props: { value: string; report: (value: string) => void }): null {
		useReportOnChange(props.value, props.report);
		return null;
	}

	it('reports transitions, never the value the component mounted with', () => {
		const report = vi.fn();
		act(() => root.render(<Probe value='idle' report={report} />));
		expect(report).not.toHaveBeenCalled();
		act(() => root.render(<Probe value='saving' report={report} />));
		act(() => root.render(<Probe value='saving' report={report} />));
		act(() => root.render(<Probe value='saved' report={report} />));
		expect(report.mock.calls).toStrictEqual([['saving'], ['saved']]);
	});
});

describe('usePropertyHandlers host events', () => {
	function mount(extra: Partial<UsePropertyHandlersInput>) {
		let handlers!: ReturnType<typeof usePropertyHandlers>;
		const markDirty = vi.fn();
		const input = {
			slides: [],
			activeSlideIndex: 0,
			canvasSize: { width: 960, height: 540 },
			setContent: vi.fn(),
			setPresentationProperties: vi.fn(),
			setCoreProperties: vi.fn(),
			setAppProperties: vi.fn(),
			setCustomProperties: vi.fn(),
			setSlides: vi.fn(),
			setIsDirty: vi.fn(),
			ops: { updateSlides: vi.fn() },
			history: { markDirty },
			...extra,
		} as unknown as UsePropertyHandlersInput;
		function Probe(): null {
			handlers = usePropertyHandlers(input);
			return null;
		}
		act(() => root.render(<Probe />));
		return { handlers: () => handlers, markDirty };
	}

	it('onPropertiesChange receives the fields the properties dialog changed', () => {
		const onPropertiesChange = vi.fn();
		const { handlers, markDirty } = mount({ onPropertiesChange });
		const updates: Partial<PptxCoreProperties> = { title: 'Quarterly review' };
		act(() => handlers().handleUpdateCoreProperties(updates));
		expect(onPropertiesChange).toHaveBeenCalledExactlyOnceWith(updates);
		expect(markDirty).toHaveBeenCalledOnce();
	});

	it('onNotesUpdate receives the new notes text', () => {
		const onNotesUpdate = vi.fn();
		const { handlers } = mount({ onNotesUpdate });
		act(() => handlers().handleUpdateNotes('Say hello', []));
		expect(onNotesUpdate).toHaveBeenCalledExactlyOnceWith('Say hello');
	});

	it('works without the host callbacks', () => {
		const { handlers } = mount({});
		expect(() => {
			act(() => {
				handlers().handleUpdateCoreProperties({ title: 'x' });
				handlers().handleUpdateNotes('y');
			});
		}).not.toThrow();
	});
});
