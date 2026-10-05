import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';
import { OFFICE_UI_TAGS, registerOfficeUi } from './index.js';
import { buildManifest } from './manifest.js';

const FILE = resolve(process.cwd(), 'custom-elements.json');

beforeAll(() => registerOfficeUi());

describe('custom-elements.json', () => {
	it('lists every element once, with its attributes, properties and events', () => {
		const manifest = buildManifest();
		const tags = manifest.modules.flatMap((module) =>
			module.declarations.map((declaration) => declaration.tagName),
		);
		expect(tags.sort()).toEqual([...OFFICE_UI_TAGS].sort());
		const button = manifest.modules
			.flatMap((module) => module.declarations)
			.find((declaration) => declaration.tagName === 'office-ui-button')!;
		expect(button.attributes.map((attribute) => attribute.name)).toContain('label');
		expect(button.members.map((member) => member.name)).toContain('disabled');
		expect(button.events.map((event) => event.name)).toContain('office-command');
		expect(button.description).toContain('Command button');
	});

	// Run `bun run gen:manifest` after changing an element; the file ships with the package.
	it('is up to date', () => {
		const next = `${JSON.stringify(buildManifest(), null, '\t')}\n`;
		if (process.env.UPDATE_MANIFEST) writeFileSync(FILE, next);
		expect(readFileSync(FILE, 'utf8').replaceAll('\r\n', '\n')).toBe(next);
	});
});
