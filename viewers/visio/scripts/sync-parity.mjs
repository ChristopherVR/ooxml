import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const file = resolve(root, 'docs/parity.md');
const html = new JSDOM(readFileSync(resolve(root, 'docs/parity.html'), 'utf8'));
const normalize = (value) => value.replace(/\s+/g, ' ').trim();
const rows = [...html.window.document.querySelectorAll('table tr')].map((row) =>
	[...row.cells].map((cell) => normalize(cell.textContent)),
);
if (rows.length < 2 || rows.some((row) => row.length !== 5))
	throw new Error('Expected one five-column capability table.');
const markdown = readFileSync(file, 'utf8');
const lines = markdown.split('\n');
const first = lines.findIndex((line) => line.startsWith('|'));
const last = lines.findLastIndex((line) => line.startsWith('|'));
if (first < 0 || last <= first) throw new Error('Markdown capability table is missing.');
const existing = lines
	.slice(first, last + 1)
	.filter((line) => line.startsWith('|'))
	.filter((line) => !/^\|[\s:|-]+\|$/.test(line))
	.map((line) => line.slice(1, -1).split('|').map(normalize));
if (process.argv.includes('--check')) {
	if (JSON.stringify(existing) !== JSON.stringify(rows))
		throw new Error(
			'Capability ledgers differ. Update docs/parity.html, then run npm run docs:sync.',
		);
	console.log(`Capability ledgers agree: ${rows.length - 1} entries.`);
} else {
	if (rows.some((row) => row.some((cell) => cell.includes('|'))))
		throw new Error('Literal pipes require explicit Markdown escaping support.');
	const table = [rows[0], Array(5).fill('---'), ...rows.slice(1)].map(
		(row) => `| ${row.join(' | ')} |`,
	);
	lines.splice(first, last - first + 1, ...table);
	writeFileSync(file, lines.join('\n'));
	console.log(`Synchronized ${rows.length - 1} capability entries.`);
}
html.window.close();
