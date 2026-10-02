// @vitest-environment node
import { expect, it } from 'vitest';

it('can be imported and registered during server rendering', async () => {
	const module = await import('./index');
	expect(() => module.defineXlsxEditor()).not.toThrow();
	expect(module.normalizeEditorLocale('de-AT')).toBe('de');
	expect(module.XLSX_EDITOR_EVENTS).toContain('file-command');
}, 30_000);
