// @vitest-environment happy-dom
/**
 * The export progress modal is the shared `pptx-ui-export-progress`: a modal dialog named by its
 * title with a determinate progress bar, and Cancel handed to the host.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { ExportProgressModal } from './ExportProgressModal';

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

const inner = (selector: string) =>
	container.querySelector('pptx-ui-export-progress')?.shadowRoot?.querySelector(selector) ?? null;

describe('exportProgressModal', () => {
	it('shows a modal dialog with a progress bar while open', () => {
		act(() =>
			root.render(
				<ExportProgressModal
					isOpen
					title='Export as PDF'
					progress={57.6}
					statusMessage='Rendering slide 4 of 7...'
					onCancel={() => {}}
				/>,
			),
		);
		const dialog = inner('[role="dialog"]')!;
		expect(dialog.getAttribute('aria-modal')).toBe('true');
		expect(inner('#title')!.textContent).toBe('Export as PDF');
		expect(inner('[role="progressbar"]')!.getAttribute('aria-valuenow')).toBe('58');
		expect(inner('#status')!.textContent).toBe('Rendering slide 4 of 7...');
	});

	it('hands Cancel to the host and draws nothing while closed', () => {
		const cancel = vi.fn();
		const view = (isOpen: boolean) => (
			<ExportProgressModal isOpen={isOpen} title='Export' progress={10} onCancel={cancel} />
		);
		act(() => root.render(view(true)));
		act(() => (inner('[data-testid="pptx-export-progress-cancel"]') as HTMLElement).click());
		expect(cancel).toHaveBeenCalledOnce();
		act(() => root.render(view(false)));
		expect(inner('[role="dialog"]')).toBeNull();
		expect(container.querySelector('pptx-ui-export-progress')).not.toBeNull();
	});
});
