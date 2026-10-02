import { describe, it, expect } from 'vitest';
import { createVsdxFixture } from '../tests/fixture.mjs';
import { mountViewer } from './binding.js';

describe('end-to-end import in DOM runtime', () => {
	it('opens a synthetic VSDX through package parser, controller and SVG renderer', async () => {
		const viewer = mountViewer(document.createElement('div'));
		await viewer.load(await createVsdxFixture('<script>literal text</script>'));
		expect(viewer.element.document?.pages[0]?.name).toBe('Imported page');
		expect(viewer.element.shadowRoot?.querySelector('svg text')?.textContent).toContain(
			'<script>literal text</script>',
		);
		expect(viewer.element.shadowRoot?.querySelector('script')).toBeNull();
		expect(
			viewer.element.document?.diagnostics.some((note) => note.code === 'cached-values-only'),
		).toBe(true);
		viewer.destroy();
	});
});
