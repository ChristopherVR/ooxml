import { afterEach, describe, expect, it } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { mountViewer } from './binding';

afterEach(() => document.body.replaceChildren());

describe('Visio ribbon collapse', () => {
	it('supports Ctrl+F1 and double-click through the shared ribbon', async () => {
		const host = document.body.appendChild(document.createElement('div'));
		const viewer = mountViewer(host, { document: structuredClone(demoDocument) });
		try {
			const ribbon = viewer.element.shadowRoot!.querySelector<
				HTMLElement & { updateComplete: Promise<unknown> }
			>('office-ui-ribbon')!;
			await ribbon.updateComplete;
			const home = ribbon.shadowRoot!.querySelector<HTMLButtonElement>('[role="tab"]')!;
			home.dispatchEvent(
				new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true, bubbles: true, composed: true }),
			);
			expect(ribbon.hasAttribute('collapsed')).toBe(true);
			home.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));
			expect(ribbon.hasAttribute('collapsed')).toBe(false);
		} finally {
			viewer.destroy();
		}
	});
});
