// @vitest-environment node
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

// The framework libraries are not installed in this package, so the component bindings are
// checked at the source: every one routes props, events and the host class through bindTeams.
const bindings = {
	react: 'react/src/index.ts',
	vue: 'vue/src/index.ts',
	solid: 'solid/src/index.ts',
	svelte: 'svelte/src/Teams.svelte',
	angular: 'angular/src/index.ts',
	vanilla: 'vanilla/src/index.ts',
} as const;
const classProp = {
	react: /className/,
	vue: /fall through/,
	solid: /props\.class/,
	svelte: /props\.class/,
	angular: /@Input\(\) className/,
	vanilla: /TeamsElementProps/,
} as const;

describe.each(Object.entries(bindings))('%s binding', (name, file) => {
	const source = readFileSync(resolve(process.cwd(), 'packages', file), 'utf8');
	it('uses the shared bindTeams wiring instead of its own copy', () => {
		expect(source).toContain('bindTeams(');
		expect(source).not.toContain('applyTeamsProps(');
		expect(source).not.toContain('listenTeamsEvents(');
	});
	it('accepts a host class', () => {
		expect(source).toMatch(classProp[name as keyof typeof classProp]);
	});
});
