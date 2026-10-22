#!/usr/bin/env node
/**
 * announce-discord.mjs: post one Discord message per release run.
 *
 * A release run can tag many packages at once (the core, the UI, the MCP server and every viewer
 * package), so GitHub's own "release published" webhook would post a message per package. This
 * reads the release plan instead and sends a single summary: the packages that shipped with their
 * release links, then the notes of the highlighted packages (core, ui, mcp), and whether npm
 * publishing succeeded.
 *
 * Usage:
 *   node scripts/announce-discord.mjs --plan release-plan.json --status success|failure \
 *     [--run-url <url>] [--dry-run]
 *
 * The webhook URL is read from DISCORD_WEBHOOK_URL. When it is unset (a fork, or the secret is not
 * configured yet) the script logs that and exits 0, so announcing never fails a release. Mentions
 * are disabled in the payload: release notes come from commit subjects and must not ping anyone.
 * Needs the `gh` CLI (GH_TOKEN in CI) to read each release's body and URL.
 */
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

/** Plan keys whose notes are quoted in the message; every other package is only listed. */
export const HIGHLIGHT_KEYS = ['core', 'ui', 'mcp'];

const COLOR = { success: 0x2ecc71, failure: 0xe74c3c };
/** Discord caps one message at 6000 characters across all embeds, and 4096 per description. */
const TOTAL_BUDGET = 5600;
const SUMMARY_MAX = 2000;
const NOTES_MAX = 1200;

/** Cut `text` to `max` characters on a line boundary, adding an ellipsis when it was shortened. */
export function truncate(text, max) {
	if (text.length <= max) return text;
	const cut = text.slice(0, max - 1);
	const nl = cut.lastIndexOf('\n');
	return `${(nl > max / 2 ? cut.slice(0, nl) : cut).trimEnd()}…`;
}

/** The release body without the "superseded releases are pruned" footer the workflow appends. */
export function cleanNotes(body) {
	return (body ?? '').split(/^---$/mu)[0].trim();
}

/**
 * @param {object} input
 * @param {string} input.repo `owner/name`
 * @param {'success' | 'failure'} input.status result of the npm publish
 * @param {string} [input.runUrl] the workflow run
 * @param {{key: string, npm: string, version: string, url: string, notes: string}[]} input.releases
 */
export function buildPayload({ repo, status, runUrl, releases }) {
	const ok = status === 'success';
	const list = releases.map((r) => `• **${r.npm}** [${r.version}](${r.url})`).join('\n');
	const summary = {
		title: ok
			? `Released ${releases.length} package${releases.length === 1 ? '' : 's'}`
			: 'Release tagged, but npm publishing did not finish',
		url: `https://github.com/${repo}/releases`,
		color: ok ? COLOR.success : COLOR.failure,
		description: truncate(
			ok
				? list
				: `${list}\n\nCheck the [workflow run](${runUrl ?? `https://github.com/${repo}/actions`}).`,
			SUMMARY_MAX,
		),
		footer: { text: ok ? 'Published to npm with provenance' : 'Re-run with a tag dispatch' },
	};
	const embeds = [summary];
	let budget = TOTAL_BUDGET - summary.title.length - summary.description.length;
	for (const key of HIGHLIGHT_KEYS) {
		const release = releases.find((r) => r.key === key);
		if (!release || !release.notes) continue;
		const title = `${release.npm} ${release.version}`;
		const description = truncate(release.notes, Math.min(NOTES_MAX, budget - title.length));
		if (description.length < 20) break;
		embeds.push({ title, url: release.url, color: summary.color, description });
		budget -= title.length + description.length;
	}
	return {
		username: 'ooxml releases',
		allowed_mentions: { parse: [] },
		embeds,
	};
}

/** Releases of the plan that were cut this run, in publish order, with their GitHub release data. */
export function releasedPackages(plan) {
	return plan.order.map((key) => ({ key, ...plan.packages[key] })).filter((p) => p.release);
}

function fetchRelease(tag, repo, gh = execFileSync) {
	const fallback = {
		url: `https://github.com/${repo}/releases/tag/${encodeURIComponent(tag)}`,
		notes: '',
	};
	try {
		const out = gh('gh', ['release', 'view', tag, '--repo', repo, '--json', 'body,url'], {
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'ignore'],
		});
		const data = JSON.parse(out);
		return { url: data.url || fallback.url, notes: cleanNotes(data.body) };
	} catch {
		return fallback;
	}
}

function parseArgs(argv) {
	const args = { plan: 'release-plan.json', status: 'success', runUrl: undefined, dryRun: false };
	for (let i = 0; i < argv.length; i++) {
		const a = argv[i];
		if (a === '--plan') args.plan = argv[++i];
		else if (a === '--status') args.status = argv[++i];
		else if (a === '--run-url') args.runUrl = argv[++i];
		else if (a === '--dry-run') args.dryRun = true;
		else throw new Error(`Unknown argument ${a}`);
	}
	if (args.status !== 'success' && args.status !== 'failure') {
		throw new Error(`--status must be success or failure, got ${args.status}`);
	}
	return args;
}

async function post(webhook, payload) {
	for (let attempt = 1; attempt <= 3; attempt++) {
		const res = await fetch(webhook, {
			method: 'POST',
			headers: { 'content-type': 'application/json' },
			body: JSON.stringify(payload),
		});
		if (res.ok) return;
		if (res.status === 429 && attempt < 3) {
			const wait = Number((await res.json().catch(() => ({}))).retry_after ?? 1);
			await new Promise((resolve) => setTimeout(resolve, Math.min(wait, 10) * 1000));
			continue;
		}
		throw new Error(`Discord answered HTTP ${res.status}`);
	}
}

async function main() {
	const args = parseArgs(process.argv.slice(2));
	const repo = process.env.GITHUB_REPOSITORY || 'ChristopherVR/ooxml';
	const plan = JSON.parse(readFileSync(args.plan, 'utf8'));
	const releases = releasedPackages(plan).map((p) => ({
		key: p.key,
		npm: p.npm,
		version: p.version,
		...fetchRelease(p.tag, repo),
	}));
	if (releases.length === 0) {
		console.log('No packages were released; nothing to announce.');
		return;
	}
	const payload = buildPayload({ repo, status: args.status, runUrl: args.runUrl, releases });
	if (args.dryRun) {
		console.log(JSON.stringify(payload, null, 2));
		return;
	}
	const webhook = process.env.DISCORD_WEBHOOK_URL;
	if (!webhook) {
		console.log('DISCORD_WEBHOOK_URL is not set; skipping the Discord announcement.');
		return;
	}
	await post(webhook, payload);
	console.log(`Announced ${releases.length} release(s) on Discord.`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
	main().catch((error) => {
		// Announcing is best effort: a Discord outage must not turn a finished release red.
		console.log(`::warning::Discord announcement failed: ${error.message}`);
	});
}
