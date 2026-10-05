// Regenerates custom-elements.json from the element classes (the file ships with the package and a
// test fails while it is out of date).
import { execFileSync } from 'node:child_process';

execFileSync('bunx', ['vitest', 'run', 'src/manifest.test.ts'], {
	stdio: 'inherit',
	env: { ...process.env, UPDATE_MANIFEST: '1' },
	shell: process.platform === 'win32',
});
