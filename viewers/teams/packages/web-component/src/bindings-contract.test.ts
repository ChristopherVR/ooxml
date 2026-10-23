// @vitest-environment node
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { angularOutputName, optionNames, propKeys } from 'ooxml-ui/bindings';
import { describe, expect, it } from 'vitest';
import { TEAMS_CONTRACT } from './index';

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

// The props and events of every binding against TEAMS_CONTRACT. The bindings that name their
// props one by one (Vue, Angular) are read at the source; the others pass every prop through
// pickTeamsProps, which reads the contract.
describe('contract parity', () => {
	const read = (file: string): string =>
		readFileSync(resolve(process.cwd(), 'packages', file), 'utf8');
	const sorted = (values: Iterable<string>) => [...values].sort();
	const options = optionNames(TEAMS_CONTRACT);
	it('lists the shared props, including locale', () => {
		expect(sorted(propKeys(TEAMS_CONTRACT))).toEqual(
			sorted([
				'config',
				'embeds',
				'locale',
				'openers',
				'uploadFile',
				'userId',
				'userName',
				'workspaceId',
			]),
		);
	});
	it('Vue declares one prop per key and emits one event per handler', () => {
		const source = read('vue/src/index.ts');
		const props = /props: \{([\s\S]*?)\n\t\},/u.exec(source)?.[1] ?? '';
		const names = [...props.matchAll(/^\t\t(\w+):/gmu)].map((match) => match[1]!);
		expect(sorted(names)).toEqual(sorted(propKeys(TEAMS_CONTRACT)));
		const emits = /emits: \[([^\]]*)\]/u.exec(source)?.[1] ?? '';
		const kebab = (option: string) =>
			option.slice(2).replace(/[A-Z]/gu, (c, at: number) => (at ? '-' : '') + c.toLowerCase());
		expect(sorted(emits.match(/[\w-]+/gu) ?? [])).toEqual(sorted(options.map(kebab)));
	});
	it('Angular has one input per key (plus className) and one output per handler', () => {
		const source = read('angular/src/index.ts');
		const inputs = [...source.matchAll(/@Input\(\) (\w+)/gu)].map((match) => match[1]!);
		expect(sorted(inputs)).toEqual(sorted([...propKeys(TEAMS_CONTRACT), 'className']));
		const outputs = [...source.matchAll(/@Output\(\) (\w+)/gu)].map((match) => match[1]!);
		expect(sorted(outputs)).toEqual(sorted(options.map(angularOutputName)));
	});
});
