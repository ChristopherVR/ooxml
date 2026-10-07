import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { localCoreAliases } from './local-core-aliases';

const root = fileURLToPath(new URL('../', import.meta.url));
const checkout = resolve(root, '../..');
afterEach(() => vi.unstubAllEnvs());

describe('local source aliases', () => {
	it('leaves published package resolution in place when not opted in', () => {
		vi.stubEnv('OOXML_CORE_SRC', '');
		expect(localCoreAliases(root)).toEqual([]);
	});

	it('resolves the shared editor and nested core entries to existing source entry points', () => {
		vi.stubEnv('OOXML_CORE_SRC', checkout);
		const aliases = localCoreAliases(root);
		const mapped = (specifier: string) => {
			const alias = aliases.find(({ find }) => find.test(specifier));
			return alias && resolve(specifier.replace(alias.find, alias.replacement));
		};
		expect(mapped('ooxml-ui/xlsx')).toBe(resolve(checkout, 'src/ui/src/xlsx/index.ts'));
		expect(mapped('ooxml-ui/controls')).toBe(resolve(checkout, 'src/ui/src/controls.ts'));
		expect(mapped('ooxml-core/xlsx/ui')).toBe(resolve(checkout, 'src/core/xlsx/ui/index.ts'));
		expect(mapped('ooxml-core/xlsx/load')).toBe(resolve(checkout, 'src/core/xlsx/load/index.ts'));
	});
});
