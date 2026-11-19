// @vitest-environment node
import { renderToString } from 'solid-js/web';
import { describe, expect, it } from 'vitest';
import { Teams } from './index';

describe('Teams on the server', () => {
	it('renders the empty element instead of touching the DOM', () => {
		const html = renderToString(() => Teams({ class: 'workspace' }) as unknown as string);
		expect(html).toMatch(/^<teams-app [^>]*class="workspace ?"[^>]*><\/teams-app>$/);
		expect(html).toContain('data-hk=');
	});
});
