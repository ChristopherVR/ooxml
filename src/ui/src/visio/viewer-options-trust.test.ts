import { afterEach, describe, expect, it, vi } from 'vitest';
import type { VisioDocument } from 'ooxml-core/visio';
import { demoDocument } from 'ooxml-core/visio/ui';
import { resetWebfontSessionCache } from '../fonts/webfonts';
import { mountViewer } from './binding';
import { readBehaviour } from './viewer-behaviour';
import { visioOptionCategories } from './viewer-options-categories';

const links = () => document.head.querySelectorAll('link[id^="visio-viewer-webfonts"]');
const calibri = (): VisioDocument => {
	const model = structuredClone(demoDocument);
	const shape = model.pages[0]!.shapes.find((item) => item.text.plainText)!;
	shape.text.fontFamily = 'Calibri';
	for (const run of shape.text.runs) run.fontFamily = 'Calibri';
	return model;
};

afterEach(() => {
	for (const link of document.head.querySelectorAll('link[id^="visio-viewer-webfonts"]'))
		link.remove();
	document.body.replaceChildren();
	localStorage.clear();
	resetWebfontSessionCache();
});

function setup(options: Parameters<typeof mountViewer>[1] = {}) {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: calibri(), ...options });
	const root = viewer.element.shadowRoot!;
	const dialog = root.querySelector<
		HTMLElement & { category: string; values: Record<string, unknown>; categories: unknown }
	>('office-ui-options-dialog')!;
	const open = () => {
		const file = root
			.querySelector('office-ui-ribbon')!
			.shadowRoot!.querySelector<HTMLElement>('.file')!;
		file.click();
		root
			.querySelector('office-ui-backstage')!
			.shadowRoot!.querySelector<HTMLButtonElement>('[data-backstage-item="options"]')!
			.click();
		dialog.category = 'trust';
	};
	const toggle = () =>
		dialog.shadowRoot!.querySelector<HTMLInputElement>('[data-key="allowFontDownloads"] input')!;
	const ok = () =>
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
	return { viewer, dialog, open, toggle, ok };
}

describe('Visio Trust Center', () => {
	it('has one font download setting, off by default, and requests nothing', async () => {
		const { viewer, open, toggle } = setup();
		open();
		expect(toggle().checked).toBe(false);
		expect(toggle().disabled).toBe(false);
		expect(readBehaviour().allowFontDownloads).toBe(false);
		await new Promise((resolve) => setTimeout(resolve, 0));
		expect(links()).toHaveLength(0);
		viewer.destroy();
	});

	it('remembers the choice and turns the download on and off for the open drawing', async () => {
		const { viewer, open, toggle, ok } = setup();
		open();
		toggle().click();
		ok();
		expect(JSON.parse(localStorage.getItem('ooxml-visio-options')!)).toMatchObject({
			allowFontDownloads: true,
		});
		await vi.waitFor(() => expect(links()).toHaveLength(1));
		const href = links()[0]!.getAttribute('href')!;
		expect(href.startsWith('https://fonts.googleapis.com/css2?')).toBe(true);
		expect(href).toContain('family=Carlito');
		open();
		expect(toggle().checked).toBe(true);
		toggle().click();
		ok();
		expect(links()).toHaveLength(0);
		expect(readBehaviour().allowFontDownloads).toBe(false);
		viewer.destroy();
	});

	it('starts with the saved choice', async () => {
		localStorage.setItem('ooxml-visio-options', JSON.stringify({ allowFontDownloads: true }));
		const { viewer } = setup();
		await vi.waitFor(() => expect(links()).toHaveLength(1));
		viewer.destroy();
	});

	it('shows a host decision read-only and keeps the reader choice underneath', async () => {
		localStorage.setItem('ooxml-visio-options', JSON.stringify({ allowFontDownloads: true }));
		const { viewer, open, toggle, ok } = setup();
		viewer.element.googleFonts = false;
		expect(links()).toHaveLength(0);
		open();
		expect(toggle().checked).toBe(false);
		expect(toggle().disabled).toBe(true);
		expect(toggle().closest('[title]')?.getAttribute('title')).toBe(
			'This setting is managed by your organization',
		);
		ok();
		expect(readBehaviour().allowFontDownloads).toBe(true);
		viewer.element.googleFonts = { baseUrl: 'https://fonts.example.test/css2' };
		await vi.waitFor(() => expect(links()).toHaveLength(1));
		expect(links()[0]!.getAttribute('href')).toContain('https://fonts.example.test/css2?');
		open();
		expect(toggle().checked).toBe(true);
		expect(toggle().disabled).toBe(true);
		// Handing the decision back unlocks the toggle again.
		viewer.element.googleFonts = undefined;
		open();
		expect(toggle().disabled).toBe(false);
		viewer.destroy();
	});

	it('translates the Trust Center and its lock reason', () => {
		const trust = visioOptionCategories('fr', true).find((category) => category.id === 'trust')!;
		expect(trust.label).toBe('Centre de gestion de la confidentialité');
		const control = trust.sections[0]!.controls[0]!;
		expect(control.disabled).toBe('Ce paramètre est géré par votre organisation');
		expect(control.info).toContain('Visio');
		expect(control.info).not.toContain('PowerPoint');
	});
});
