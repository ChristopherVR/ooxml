import { readFileSync } from 'node:fs';
import { compile } from 'svelte/compiler';
import { describe, expect, it } from 'vitest';
import { navigationControls, statusControls, toolbarControls } from './building-blocks';

describe('Svelte building-block components', () => {
	it.each(['Controls', 'Toolbar', 'PageNavigation', 'StatusBar', 'DocumentCanvas'])(
		'%s compiles for the client',
		(name) => {
			const source = readFileSync(new URL(`./${name}.svelte`, import.meta.url), 'utf8');
			expect(() =>
				compile(source, { filename: `${name}.svelte`, generate: 'client' }),
			).not.toThrow();
		},
	);

	it('re-exports the control descriptors the components draw', () => {
		for (const fn of [toolbarControls, navigationControls, statusControls])
			expect(fn).toBeTypeOf('function');
	});
});
