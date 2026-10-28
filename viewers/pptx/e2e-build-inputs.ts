import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';

export const workspaceRoot = resolve(__dirname, '../..');
export const demoDir = (binding: string): string =>
	join(workspaceRoot, 'demos/pptx', `demo-${binding}`);

/** Actual bytes include uncommitted edits. Stamps are private and never packaged. */
export function demoInputs(): string {
	const hash = createHash('sha256');
	const visit = (path: string): void => {
		if (!existsSync(path)) throw new Error(`Missing browser build input: ${path}`);
		for (const item of readdirSync(path, { withFileTypes: true }).sort((a, b) =>
			a.name.localeCompare(b.name),
		)) {
			if (['node_modules', '.git', 'dist', '.e2e-servers'].includes(item.name)) continue;
			const child = join(path, item.name);
			if (item.isDirectory()) visit(child);
			else if (!/\.(?:map|md)$|\.(?:test|spec)\.[cm]?[jt]sx?$/.test(item.name)) {
				hash.update(child.slice(workspaceRoot.length));
				hash.update(readFileSync(child));
			}
		}
	};
	for (const file of [
		'bun.lock',
		'package.json',
		'viewers/pptx/e2e-build.ts',
		'viewers/pptx/e2e-build-inputs.ts',
	])
		hash.update(readFileSync(join(workspaceRoot, file)));
	for (const dir of [
		'src/core/dist',
		'src/ui/dist',
		'viewers/pptx/packages',
		'demos/pptx',
		'e2e/pptx/fixtures',
	])
		visit(join(workspaceRoot, dir));
	// Angular and other dist consumers must invalidate demo output after a package rebuild.
	for (const binding of readdirSync(join(workspaceRoot, 'viewers/pptx/packages')).sort()) {
		const dist = join(workspaceRoot, 'viewers/pptx/packages', binding, 'dist');
		if (existsSync(dist)) visit(dist);
	}
	return hash.digest('hex');
}

export function demoStamp(binding: string): string {
	return join(demoDir(binding), 'node_modules/.cache/ooxml-e2e-build');
}

export function demoIsFresh(binding: string, hash: string): boolean {
	return (
		existsSync(join(demoDir(binding), 'dist/index.html')) &&
		existsSync(demoStamp(binding)) &&
		readFileSync(demoStamp(binding), 'utf8') === hash
	);
}
