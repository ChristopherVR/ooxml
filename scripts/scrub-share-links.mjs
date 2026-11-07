/**
 * Removes AI chat share links from a commit message (AGENTS.md: "Never include an AI chat share
 * link in a commit, PR, changelog, comment or doc"). Run as a `commit-msg` hook, which may rewrite
 * the message file before git records it:
 *
 *   node scripts/scrub-share-links.mjs <message-file>
 *
 * A trailer that carries a link (`Claude-Session: https://...`) and a line that is nothing but a
 * link are dropped whole; a link inside prose is cut out of its line. Each removal is reported on
 * stderr so the author sees what changed. The commit always goes ahead.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

/** Hosts and paths of AI chat sessions and their share pages. */
const SHARE_LINK_RE =
	/https?:\/\/(?:(?:www\.)?claude\.ai\/(?:code\/session|share|chat)|(?:www\.)?chatgpt\.com\/(?:share|c)|chat\.openai\.com\/(?:share|c)|gemini\.google\.com\/(?:share|app)|g\.co\/gemini\/share|copilot\.microsoft\.com\/(?:shares|chats))\/?[^\s)>\]]*/giu;

const TRAILER_RE = /^[A-Za-z][A-Za-z0-9-]*:\s/u;

/**
 * @param {string} message
 * @returns {{ message: string, removed: string[] }}
 */
export function scrubShareLinks(message) {
	/** @type {string[]} */
	const removed = [];
	/** @type {string[]} */
	const kept = [];
	for (const line of message.split('\n')) {
		// git strips `#` comment lines itself; leave them, the template lives there.
		const links = line.startsWith('#') ? null : line.match(SHARE_LINK_RE);
		if (!links) {
			kept.push(line);
			continue;
		}
		removed.push(...links);
		const rest = line.replace(SHARE_LINK_RE, '').replace(TRAILER_RE, '').trim();
		if (/^[-*:(<\[\]>)\s]*$/u.test(rest)) continue;
		kept.push(
			line
				.replace(SHARE_LINK_RE, '')
				.replace(/[ \t]{2,}/gu, ' ')
				.trimEnd(),
		);
	}
	if (removed.length === 0) return { message, removed };
	// Collapse the blank lines a removed trailer can leave at the end of the message.
	const text = kept
		.join('\n')
		.replace(/\n{3,}/gu, '\n\n')
		.replace(/\s+$/u, '');
	return { message: `${text}\n`, removed };
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
	const file = process.argv[2];
	if (!file) {
		console.error('usage: node scripts/scrub-share-links.mjs <commit-message-file>');
		process.exit(2);
	}
	const { message, removed } = scrubShareLinks(readFileSync(file, 'utf8'));
	if (removed.length > 0) {
		writeFileSync(file, message);
		for (const link of removed) console.error(`commit-msg: removed AI chat share link ${link}`);
	}
}
