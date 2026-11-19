import { readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createHash } from 'node:crypto';

/**
 * The app page's stylesheets, in cascade order. `site/` keeps them split by surface; a build
 * serves them as one content-hashed file, so the browser finds every rule with the page instead
 * of discovering `workspace.css`'s `@import`s one round trip later.
 */
const SHEETS = ['styles.css', 'themes.css', 'workspace.css'];
/** The docs hub and the not-found page link the palette on its own, so it stays in the build. */
const SHARED = new Set(['themes.css']);
const IMPORT = /^@import url\('\.\/([\w-]+\.css)'\);$/gm;
const LINK = /[ \t]*<link rel="stylesheet" href="([\w-]+\.css)" \/>\n/g;

/** `name` with its `@import`s replaced, in place, by the sheets they name (cascade order kept). */
async function inline(dir, name, used) {
	if (used.has(name)) throw new Error(`${name} is imported twice`);
	used.add(name);
	const css = await readFile(join(dir, name), 'utf8');
	const parts = [];
	let last = 0;
	for (const match of css.matchAll(IMPORT)) {
		parts.push(css.slice(last, match.index), await inline(dir, match[1], used));
		last = match.index + match[0].length;
	}
	parts.push(css.slice(last));
	return parts.join('');
}

/**
 * Writes `shell-<hash>.css` into the build `out` from the copied sheets, points `index.html` at
 * it instead of the separate sheets and removes the sheets it absorbed. Returns the file name.
 */
export async function bundleShellCss(out) {
	const page = join(out, 'index.html');
	const html = await readFile(page, 'utf8');
	const links = [...html.matchAll(LINK)].map((match) => match[1]).filter((h) => SHEETS.includes(h));
	if (links.join() !== SHEETS.join()) throw new Error(`index.html must link ${SHEETS.join(', ')}`);
	const used = new Set();
	const css = [];
	for (const sheet of SHEETS) css.push(await inline(out, sheet, used));
	const text = css.join('\n');
	const name = `shell-${createHash('sha256').update(text).digest('hex').slice(0, 12)}.css`;
	await writeFile(join(out, name), text);
	// The first sheet's link becomes the bundle's; the others go.
	await writeFile(
		page,
		html.replace(LINK, (match, href) =>
			!SHEETS.includes(href) ? match : href === SHEETS[0] ? match.replace(href, name) : '',
		),
	);
	for (const file of used) if (!SHARED.has(file)) await rm(join(out, file));
	return name;
}
