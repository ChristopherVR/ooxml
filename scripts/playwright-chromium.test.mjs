import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, it } from 'node:test';
import { findInstalledChromium, resolveChromiumExecutable } from './playwright-chromium.mjs';

function fakeCache() {
	const root = mkdtempSync(join(tmpdir(), 'pw-'));
	for (const [dir, exe] of [
		['chromium_headless_shell-1100', 'headless_shell'],
		['chromium_headless_shell-1194', 'headless_shell'],
		['chromium-1200', 'chrome'],
	]) {
		for (const [platformDir, filename] of [
			['chrome-linux', exe],
			['chrome-win', `${exe}.exe`],
			['chrome-mac', exe],
		]) {
			mkdirSync(join(root, dir, platformDir), { recursive: true });
			writeFileSync(join(root, dir, platformDir, filename), '');
		}
	}
	return root;
}

describe('chromium resolution', () => {
	it('prefers the newest headless shell', () => {
		const root = fakeCache();
		assert.equal(
			findInstalledChromium(root, 'linux'),
			join(root, 'chromium_headless_shell-1194', 'chrome-linux', 'headless_shell'),
		);
	});
	it('honours the explicit override and leaves an installed bundle alone', () => {
		const root = fakeCache();
		const bundled = join(root, 'chromium-1200', 'chrome-linux', 'chrome');
		assert.equal(
			resolveChromiumExecutable(bundled, { PLAYWRIGHT_CHROMIUM_EXECUTABLE: '/x' }),
			'/x',
		);
		assert.equal(resolveChromiumExecutable(bundled, { PLAYWRIGHT_BROWSERS_PATH: root }), undefined);
	});
	it('falls back only when the bundled executable is missing', () => {
		const root = fakeCache();
		assert.ok(
			resolveChromiumExecutable('/nope', { PLAYWRIGHT_BROWSERS_PATH: root })?.includes('1194'),
		);
	});
});
