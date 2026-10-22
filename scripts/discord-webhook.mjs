#!/usr/bin/env node
/**
 * discord-webhook.mjs: wire a Discord webhook into this repository without ever printing it.
 *
 * Copy the webhook URL in Discord (channel settings, Integrations, Webhooks, Copy Webhook URL),
 * then run one of:
 *
 *   node scripts/discord-webhook.mjs prs        # repository webhook: pull request activity
 *   node scripts/discord-webhook.mjs issues     # repository webhook: issues, comments, discussions
 *   node scripts/discord-webhook.mjs releases   # Actions secret DISCORD_WEBHOOK_RELEASES
 *   node scripts/discord-webhook.mjs ci         # Actions secret DISCORD_WEBHOOK_CI
 *
 * The URL is read from the clipboard and sent straight to `gh` on stdin: it is never echoed, put
 * in an argument list or written to disk. `prs` and `issues` use GitHub's native Discord endpoint
 * (the URL plus `/github`), and re-running one replaces its existing webhook instead of adding a
 * second. Needs the `gh` CLI with repository admin rights. See docs/discord.md.
 */
import { execFileSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

export const REPO = 'ChristopherVR/ooxml';

/** What each kind does: a native repository webhook (`events`) or an Actions secret. */
export const KINDS = {
	prs: { events: ['pull_request', 'pull_request_review', 'pull_request_review_comment'] },
	issues: { events: ['issues', 'issue_comment', 'discussion', 'discussion_comment'] },
	releases: { secret: 'DISCORD_WEBHOOK_RELEASES' },
	ci: { secret: 'DISCORD_WEBHOOK_CI' },
};

const WEBHOOK = /^https:\/\/(?:discord|discordapp)\.com\/api\/webhooks\/\d+\/[\w-]+$/u;

/** The trimmed URL when `text` is a Discord webhook URL, else null. */
export function parseWebhook(text) {
	const url = text.trim().replace(/\/github$/u, '');
	return WEBHOOK.test(url) ? url : null;
}

/** The request body for a native repository webhook that posts to Discord's GitHub endpoint. */
export function hookBody(url, events) {
	return {
		name: 'web',
		active: true,
		events,
		config: { url: `${url}/github`, content_type: 'json', insecure_ssl: '0' },
	};
}

function readClipboard() {
	if (process.platform === 'win32') {
		return execFileSync('powershell', ['-NoProfile', '-Command', 'Get-Clipboard'], {
			encoding: 'utf8',
		});
	}
	if (process.platform === 'darwin') return execFileSync('pbpaste', { encoding: 'utf8' });
	try {
		return execFileSync('wl-paste', { encoding: 'utf8' });
	} catch {
		return execFileSync('xclip', ['-selection', 'clipboard', '-o'], { encoding: 'utf8' });
	}
}

const gh = (args, input) =>
	execFileSync('gh', args, { encoding: 'utf8', input, stdio: ['pipe', 'pipe', 'inherit'] });

function sameEvents(a, b) {
	return a.length === b.length && [...a].sort().join() === [...b].sort().join();
}

function main() {
	const kind = process.argv[2];
	const spec = KINDS[kind];
	if (!spec) {
		console.error(`usage: node scripts/discord-webhook.mjs <${Object.keys(KINDS).join('|')}>`);
		process.exit(2);
	}
	const url = parseWebhook(readClipboard());
	if (!url) {
		console.error('The clipboard does not hold a Discord webhook URL. Copy one in Discord first.');
		process.exit(1);
	}
	if (spec.secret) {
		gh(['secret', 'set', spec.secret, '--repo', REPO], url);
		console.log(`Stored ${spec.secret}.`);
		return;
	}
	const body = JSON.stringify(hookBody(url, spec.events));
	const existing = JSON.parse(gh(['api', `repos/${REPO}/hooks`])).find((h) =>
		sameEvents(h.events, spec.events),
	);
	if (existing) {
		gh(['api', '-X', 'PATCH', `repos/${REPO}/hooks/${existing.id}`, '--input', '-'], body);
		console.log(`Updated the ${kind} webhook (${existing.id}).`);
	} else {
		const created = JSON.parse(
			gh(['api', '-X', 'POST', `repos/${REPO}/hooks`, '--input', '-'], body),
		);
		console.log(`Created the ${kind} webhook (${created.id}).`);
	}
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) main();
