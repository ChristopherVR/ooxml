// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { VISIO_TEMPLATES } from 'ooxml-core/visio/ui';
import { createBackstage } from './backstage';
import { fillTemplateThumbnails, templateThumbnail } from './template-thumbnails';

afterEach(() => document.body.replaceChildren());

it('renders every template page as an SVG image', async () => {
	for (const template of VISIO_TEMPLATES) {
		const source = await templateThumbnail(template.id);
		expect(source, template.id).toMatch(/^data:image\/svg\+xml;charset=utf-8,/);
		const svg = decodeURIComponent(source!.split(',', 2)[1]!);
		expect(svg).toContain('<svg');
		// Content comes from the drawing, not a stored picture: the org chart carries its text.
		if (template.id === 'organization-chart') expect(svg).toContain('Chief Executive');
		if (template.id === 'cross-functional-flowchart') expect(svg).toContain('Fulfilment');
	}
});

it('puts one thumbnail on each New tile, once', async () => {
	const backstage = createBackstage(document);
	document.body.append(backstage);
	await fillTemplateThumbnails(backstage);
	await fillTemplateThumbnails(backstage);
	const tiles = backstage.querySelectorAll('.backstage-templates .backstage-action');
	expect(tiles).toHaveLength(VISIO_TEMPLATES.length);
	for (const tile of tiles) expect(tile.querySelectorAll('img.backstage-thumb')).toHaveLength(1);
});
