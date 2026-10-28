import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';

const [input, output] = process.argv.slice(2);
if (!input || !output)
	throw new Error('usage: node merge-browser-timings.mjs <reports-dir> <output>');
const tests = {};
function visit(dir) {
	for (const item of readdirSync(dir, { withFileTypes: true }).sort((a, b) =>
		a.name.localeCompare(b.name),
	)) {
		const path = join(dir, item.name);
		if (item.isDirectory()) visit(path);
		else if (item.name.endsWith('.json')) {
			const report = JSON.parse(readFileSync(path, 'utf8'));
			if (report.version === 1) Object.assign(tests, report.tests);
		}
	}
}
visit(input);
mkdirSync(dirname(output), { recursive: true });
writeFileSync(output, JSON.stringify({ version: 1, tests }, null, 2));
console.log(`Saved timings for ${Object.keys(tests).length} tests`);
