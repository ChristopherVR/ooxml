import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';

const plan = JSON.parse(readFileSync('release-plan.json', 'utf8'));
mkdirSync('release-notes', { recursive: true });
for (const key of plan.order.filter((key) => plan.packages[key].release)) {
	const pkg = plan.packages[key];
	const notes = execFileSync(process.execPath, ['scripts/release-notes.mjs', key], {
		encoding: 'utf8',
	});
	const path = `${pkg.dir}/CHANGELOG.md`;
	let previous = '';
	try {
		previous = readFileSync(path, 'utf8');
	} catch (error) {
		if (error.code !== 'ENOENT') throw error;
	}
	writeFileSync(
		path,
		`## ${pkg.version}\n\n${notes.trimEnd()}\n${previous ? `\n${previous.trimEnd()}\n` : ''}`,
	);
	writeFileSync(`release-notes/${pkg.tag}.md`, notes);
}
