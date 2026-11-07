import { afterEach, expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import type { VisioDocument } from 'ooxml-core/visio';
import { resetWebfontSessionCache } from '../fonts/webfonts';
import { mountViewer } from './binding';

const links = () => document.head.querySelectorAll('link[id^="visio-viewer-webfonts"]');
const calibri = (): VisioDocument => {
	const model = structuredClone(demoDocument);
	const shape = model.pages[0]!.shapes.find((item) => item.text.plainText)!;
	shape.text.fontFamily = 'Calibri';
	for (const run of shape.text.runs) run.fontFamily = 'Calibri';
	return model;
};

afterEach(() => {
	for (const link of links()) link.remove();
	resetWebfontSessionCache();
});

it('requests no webfont unless googleFonts is set', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: calibri() });
	expect(viewer.element.googleFonts).toBeUndefined();
	await new Promise((resolve) => setTimeout(resolve, 0));
	expect(links()).toHaveLength(0);
	viewer.destroy();
	host.remove();
});

it('links the clone of the open drawing and redraws the page, keeping page and zoom', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: calibri() });
	viewer.element.zoom = 1.5;
	const before = viewer.element.shadowRoot!.querySelector('svg.paper');
	viewer.element.googleFonts = { baseUrl: 'https://fonts.example.test/css2' };
	await vi.waitFor(() => expect(links()).toHaveLength(1));
	expect(links()[0]!.getAttribute('href')).toContain('family=Carlito');
	await vi.waitFor(() =>
		expect(viewer.element.shadowRoot!.querySelector('svg.paper')).not.toBe(before),
	);
	expect(viewer.element.zoom).toBe(1.5);
	expect(viewer.element.pageIndex).toBe(0);
	// The painted text names the clone after the drawing's own family.
	expect(
		viewer.element.shadowRoot!.querySelector('svg.paper text')?.getAttribute('font-family'),
	).toMatch(/^"Calibri", "Carlito"/u);
	viewer.destroy();
	expect(links()).toHaveLength(0);
	host.remove();
});
