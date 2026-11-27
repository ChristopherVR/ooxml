import { afterEach, expect, it, vi } from 'vitest';
import { watchOfficeAddIns } from '../ribbon/add-in-dialog';
import { watchWordAddIns } from './add-in-mount';
import type { EditorCore } from './editor-core';

vi.mock('../ribbon/add-in-dialog', () => ({ watchOfficeAddIns: vi.fn(() => () => {}) }));

afterEach(() => vi.mocked(watchOfficeAddIns).mockClear());

it('passes the add-in dialog and pane Word labels in the editor locale', () => {
	const core = { locale: 'fr', targetView: () => undefined } as unknown as EditorCore;
	const root = document.createElement('div').attachShadow({ mode: 'open' });
	watchWordAddIns(root, core);
	const [, t, product, locale] = vi.mocked(watchOfficeAddIns).mock.calls[0] ?? [];
	expect(product).toBe('word');
	expect(t?.('Close')).toBe('Fermer');
	expect(t?.('Not a Word label')).toBe('Not a Word label');
	(core as { locale: string }).locale = 'de';
	expect(locale?.()).toBe('de');
	expect(t?.('Close')).toBe('Schließen');
});
