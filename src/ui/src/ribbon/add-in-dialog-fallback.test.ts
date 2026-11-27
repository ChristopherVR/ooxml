import { afterEach, expect, it, vi } from 'vitest';
import type { OfficeUiDialog } from '../dialog/dialog';
import { openOfficeAddIns } from './add-in-dialog';

// A failed chunk download: the dialog module cannot load, so the fallback must need no chunk.
vi.mock('../dialog/add-ins-dialog', () => {
	throw new Error('chunk failed to load');
});

afterEach(() => document.body.replaceChildren());

it('reports a failed add-in chunk in the editor locale without loading another chunk', async () => {
	const element = document.body.appendChild(document.createElement('div'));
	const root = element.attachShadow({ mode: 'open' });
	await openOfficeAddIns(root, 'myAddIns', (label) => `[${label}]`, 'word', 'fr-FR');
	const fallback = root.querySelector<OfficeUiDialog>('[data-office-add-ins-fallback]');
	expect(fallback).not.toBeNull();
	expect(fallback?.querySelector('[role="status"]')?.textContent).toBe(
		'Impossible de charger les compléments.',
	);
	expect(fallback?.heading).toBe('[Office Add-ins]');
});
