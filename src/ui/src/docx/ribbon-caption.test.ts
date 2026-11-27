import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { localizeElement } from './localization';
import { tool } from './ribbon-parts';

const words = (button: HTMLElement) =>
	[...button.querySelectorAll('.ribbon-caption-word')].map((word) => word.textContent);

it('writes a large caption as unbreakable words, so "Add-ins" never splits at its hyphen', () => {
	const button = tool(
		'Get Add-ins',
		'addIns',
		{ type: 'officeAddIns', mode: 'getAddIns' },
		{
			large: true,
		},
	);
	const caption = button.querySelector(':scope > span');
	expect(caption?.textContent).toBe('Get Add-ins');
	expect(words(button)).toEqual(['Get', 'Add-ins']);
	const css = readFileSync(new URL('./styles/ribbon-buttons.css', import.meta.url), 'utf8');
	expect(css).toMatch(/\.ribbon-caption-word \{\s*white-space: nowrap;/);
});

it('rewrites the words when the locale changes, and back', () => {
	const host = document.createElement('div');
	host.append(
		tool('Get Add-ins', 'addIns', { type: 'officeAddIns', mode: 'getAddIns' }, { large: true }),
	);
	localizeElement(host, 'de');
	expect(words(host)).toEqual(['Add-Ins', 'abrufen']);
	localizeElement(host, 'en');
	expect(words(host)).toEqual(['Get', 'Add-ins']);
});

it('keeps a small command caption as plain text', () => {
	const button = tool('Link', 'link', { type: 'link' }, { inline: true });
	expect(button.querySelector('.ribbon-caption-word')).toBeNull();
	expect(button.querySelector(':scope > span')?.textContent).toBe('Link');
});
