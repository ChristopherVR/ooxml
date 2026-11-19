// @vitest-environment node
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

describe('element definitions', () => {
	it('all go through the registry, so every class is stamped and adopts early properties', () => {
		const root = fileURLToPath(new URL('.', import.meta.url));
		const offenders: string[] = [];
		const visit = (directory: string) => {
			for (const entry of readdirSync(directory, { withFileTypes: true })) {
				const path = join(directory, entry.name);
				if (entry.isDirectory()) visit(path);
				else if (entry.name.endsWith('.ts') && !/\.test\.ts$|^registry\.ts$/.test(entry.name))
					if (/\b(customElements|registry|target)\.define\(/.test(readFileSync(path, 'utf8')))
						offenders.push(relative(root, path));
			}
		};
		visit(root);
		expect(offenders).toEqual([]);
	});
});
