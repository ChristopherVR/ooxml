import { expect, it } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { renderPage } from './render-svg';

// Visio draws PAGENAME() on a background page with the name of the foreground page using it.
it('draws a background page name field with the foreground page name', () => {
	const model = structuredClone(demoDocument);
	const front = model.pages[0]!;
	front.name = 'Orders';
	const back = structuredClone(front);
	back.id = 'bg';
	back.name = 'Background-1';
	back.isBackground = true;
	const footer = back.shapes[0]!;
	footer.id = 'footer';
	footer.text = {
		...footer.text!,
		plainText: 'Background-1',
		runs: [],
		fields: [{ start: 0, end: 12, cached: 'Background-1', formula: 'PAGENAME()' }],
	};
	back.shapes = [footer];
	front.backgroundPageId = 'bg';
	model.pages.push(back);
	const result = renderPage(model, front);
	const text = result.svg.querySelector('[data-shape-id="footer"]')?.textContent ?? '';
	expect(text).toContain('Orders');
	expect(text).not.toContain('Background-1');
	// The background page shown on its own keeps its own name.
	const alone = renderPage(model, back);
	expect(alone.svg.querySelector('[data-shape-id="footer"]')?.textContent).toContain(
		'Background-1',
	);
	result.dispose();
	alone.dispose();
});
