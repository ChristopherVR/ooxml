// The base a push plans from. `github.event.before` is only the previous push, so once push runs
// are cancelled by newer ones, a cancelled run's commits would never be checked. Instead the base
// is the newest ancestor of HEAD whose CI run succeeded: everything since it is unverified, the
// plan covers all of it, and a newer push can safely cancel an older one.
import { execFileSync } from 'node:child_process';

/** How many commits back the search for a green commit goes before the run falls back to full. */
export const GREEN_SEARCH_COMMITS = 200;
const RUN_PAGES = 5;

const defaultExec = (...args) => execFileSync('git', args, { encoding: 'utf8' }).trim();

const defaultFetchJson = async (url, headers) => {
	const response = await fetch(url, { headers });
	if (!response.ok) throw new Error(`GitHub API ${response.status} for ${url}`);
	return response.json();
};

/** SHAs of the successful `ci.yml` push or schedule runs on main, as a Set. */
export async function successfulRunShas({ repo, token, fetchJson }) {
	const shas = new Set();
	for (let page = 1; page <= RUN_PAGES; page++) {
		const url = `https://api.github.com/repos/${repo}/actions/workflows/ci.yml/runs?branch=main&status=success&per_page=100&page=${page}`;
		const body = await fetchJson(url, {
			Authorization: `Bearer ${token}`,
			Accept: 'application/vnd.github+json',
			'X-GitHub-Api-Version': '2022-11-28',
		});
		const runs = body?.workflow_runs ?? [];
		for (const run of runs)
			if (run.conclusion === 'success' && (run.event === 'push' || run.event === 'schedule'))
				shas.add(run.head_sha);
		if (runs.length < 100) break;
	}
	return shas;
}

/**
 * The base a push should plan from: `{ base, source }`. Source is `before` (the event's own base,
 * also when the lookup failed), `green` (an older green ancestor) or `none` (no green commit within
 * the search window: the base is empty, which plans a full run). `exec` runs git and
 * `fetchJson(url, headers)` calls the API; both are injectable.
 */
export async function pushBase({
	before,
	repo = process.env.GITHUB_REPOSITORY,
	token = process.env.GITHUB_TOKEN,
	exec = defaultExec,
	fetchJson = defaultFetchJson,
	limit = GREEN_SEARCH_COMMITS,
}) {
	try {
		if (!repo || !token) throw new Error('no repository or token');
		const green = await successfulRunShas({ repo, token, fetchJson });
		const head = exec('rev-parse', 'HEAD');
		const ancestors = exec('rev-list', '-n', String(limit), head).split('\n').filter(Boolean);
		const found = ancestors.find((sha) => sha !== head && green.has(sha));
		if (!found) return { base: '', source: 'none' };
		if (found === before) return { base: before, source: 'before' };
		// Only a commit older than `before` widens the plan; a newer one cannot cover its changes.
		let older = false;
		try {
			exec('merge-base', '--is-ancestor', found, before);
			older = true;
		} catch {
			// Not an ancestor of `before`.
		}
		return older ? { base: found, source: 'green' } : { base: before, source: 'before' };
	} catch {
		return { base: before, source: 'before' };
	}
}
