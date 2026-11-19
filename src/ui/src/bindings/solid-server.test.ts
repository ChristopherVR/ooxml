// @vitest-environment node
import { describe, expect, it, vi } from 'vitest';

// The workspace test setup points Solid at its browser build; a server renders with this one.
vi.mock('solid-js', () => vi.importActual('../../node_modules/solid-js/dist/server.js'));
vi.mock('solid-js/web', () => vi.importActual('../../node_modules/solid-js/web/dist/server.js'));

describe('createSolidBinding on the server', () => {
	it('renders the empty host instead of touching the DOM', async () => {
		const { renderToString } = await import('solid-js/web');
		const { createSolidBinding } = await import('./solid');
		const mount = vi.fn();
		const Editor = createSolidBinding<{ class?: string }, object, object>({
			handle: {} as never,
			mount,
			options: () => ({}),
		});
		const html = renderToString(() => Editor({ class: 'sheet' }) as unknown as string);
		expect(html).toMatch(/^<div [^>]*class="sheet ?"[^>]*><\/div>$/);
		expect(html).toContain('data-hk=');
		expect(mount).not.toHaveBeenCalled();
	});
});
