// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { shellFixture, spyCommand } from '../test-support/shell';
import { renderCellStyleGallery } from './cell-style-gallery';
import type { RenderScope } from './controls';
import { applyRibbonVisibility } from './visibility';

afterEach(() => document.body.replaceChildren());

describe('Cell Styles inline gallery', () => {
	it('paints previews, applies an inline choice and disables it in a read-only context', () => {
		const { core } = shellFixture();
		let enabled = true;
		const command = spyCommand('home.cell-styles', { enabled: () => enabled });
		core.commands.register(command);
		const scope: RenderScope = { ctx: core.ctx, doc: document, updates: [], isHidden: () => false };
		const gallery = renderCellStyleGallery(
			scope,
			{
				kind: 'gallery',
				command: command.id,
				items: () => [{ id: 'Good', label: 'Good', preview: 'background:#c6efce;color:#006100' }],
			},
			command,
		);
		core.ctx.root.append(gallery);
		gallery.setAttribute('mode', 'inline');
		for (const update of scope.updates) update();
		const tile = gallery.querySelector<HTMLButtonElement>('.strip .tile')!;
		expect(tile.querySelector('svg')!.namespaceURI).toBe('http://www.w3.org/2000/svg');
		expect(tile.querySelector('rect')!.getAttribute('fill')).toBe('rgb(198, 239, 206)');
		tile.click();
		expect(command.runs).toEqual(['Good']);
		enabled = false;
		for (const update of scope.updates) update();
		expect(gallery.querySelector<HTMLButtonElement>('.strip .tile')!.disabled).toBe(true);
	});

	it('hides the complete gallery when its command is hidden', () => {
		const { core } = shellFixture();
		const command = spyCommand('home.cell-styles');
		core.commands.register(command);
		const scope: RenderScope = { ctx: core.ctx, doc: document, updates: [], isHidden: () => false };
		const gallery = renderCellStyleGallery(
			scope,
			{
				kind: 'gallery',
				command: command.id,
				items: () => [{ id: 'Good', label: 'Good' }],
			},
			command,
		);
		const panel = document.createElement('div');
		panel.className = 'ribbon-panel';
		panel.append(gallery);
		core.ctx.root.append(panel);
		for (const update of scope.updates) update();
		applyRibbonVisibility(panel, (id) => id === command.id);
		expect(gallery.dataset.command).toBe(command.id);
		expect(gallery.hasAttribute('data-xve-hidden')).toBe(true);
	});
});
