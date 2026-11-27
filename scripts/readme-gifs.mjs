#!/usr/bin/env node
/**
 * readme-gifs.mjs: records the animated README pictures from the built suite.
 *
 * Each scene opens the suite (or one product's standalone app) in headless Chromium, waits until
 * the editor has drawn and its fonts have loaded, then plays a short scripted interaction while
 * Playwright records a video. ffmpeg cuts the loading off the front and turns the rest into a
 * palette-optimised GIF next to the README picture it replaces. The READMEs link each GIF to the
 * interactive demo on https://christophervr.github.io/ooxml/.
 *
 *   bun run build:suite                      # once, or after a UI change
 *   bun run readme:gifs                      # every scene (serves suite-dist on port 8123)
 *   node scripts/readme-gifs.mjs word excel  # only these scenes (names in readme-gifs/scenes.mjs)
 *   PORT=8130 bun run readme:gifs            # serve on another port
 *   SUITE_URL=http://localhost:8123 bun run readme:gifs   # use a preview that is already running
 *
 * Needs ffmpeg on PATH (or FFMPEG=<path>). Playwright comes from the e2e/docx package, as in the
 * other suite checks. Pass --keep to keep the raw recordings in the system temporary folder. A
 * scene's GIF should stay near 2.5 MB; the run warns when one is larger.
 */
import { spawn, spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync, statSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SCENES } from './readme-gifs/scenes.mjs';
import { cursorInit } from './readme-gifs/cursor.mjs';

const ROOT = resolve(fileURLToPath(new URL('..', import.meta.url)));
const require = createRequire(join(ROOT, 'e2e/docx/package.json'));
const { chromium } = require('@playwright/test');

const args = process.argv.slice(2);
const keep = args.includes('--keep');
const only = args.filter((a) => !a.startsWith('--'));
const unknown = only.filter((name) => !SCENES.some((s) => s.name === name));
if (unknown.length) {
	console.error(
		`Unknown scene: ${unknown.join(', ')}. Scenes: ${SCENES.map((s) => s.name).join(', ')}`,
	);
	process.exit(1);
}
const scenes = only.length ? SCENES.filter((s) => only.includes(s.name)) : SCENES;

const FFMPEG = process.env.FFMPEG ?? 'ffmpeg';
if (spawnSync(FFMPEG, ['-version'], { stdio: 'ignore' }).status !== 0) {
	console.error('ffmpeg was not found. Install it or set FFMPEG to its path.');
	process.exit(1);
}

/** Serve suite-dist with the preview server unless SUITE_URL names one already running. */
async function startServer() {
	if (process.env.SUITE_URL) return { base: process.env.SUITE_URL.replace(/\/$/, ''), stop() {} };
	if (!existsSync(join(ROOT, 'suite-dist/index.html'))) {
		console.error('suite-dist is missing: run `bun run build:suite` first.');
		process.exit(1);
	}
	const port = process.env.PORT ?? '8123';
	const child = spawn(process.execPath, [join(ROOT, 'scripts/preview-site.mjs'), port], {
		stdio: 'ignore',
	});
	const base = `http://localhost:${port}`;
	for (let i = 0; i < 100; i++) {
		const ok = await fetch(base).then(
			(r) => r.ok,
			() => false,
		);
		if (ok) return { base, stop: () => child.kill() };
		await new Promise((r) => setTimeout(r, 100));
	}
	child.kill();
	throw new Error(`The preview server did not answer on ${base}`);
}

/** Cut `trim` seconds off the front, keep `duration`, and write a two-pass palette GIF. */
function toGif(video, out, { trim, duration, width, fps }) {
	const filter =
		`fps=${fps},scale=${width}:-1:flags=lanczos,split[a][b];` +
		'[a]palettegen=max_colors=192:stats_mode=diff[p];' +
		'[b][p]paletteuse=dither=bayer:bayer_scale=4:diff_mode=rectangle';
	const result = spawnSync(
		FFMPEG,
		[
			'-y',
			'-v',
			'error',
			'-ss',
			trim.toFixed(2),
			'-t',
			duration.toFixed(2),
			'-i',
			video,
			'-filter_complex',
			filter,
			'-loop',
			'0',
			out,
		],
		{ stdio: 'inherit' },
	);
	if (result.status !== 0) throw new Error(`ffmpeg failed for ${out}`);
}

const server = await startServer();
const browser = await chromium.launch({ headless: true });
const work = mkdtempSync(join(tmpdir(), 'readme-gifs-'));
let failed = false;
try {
	for (const scene of scenes) {
		console.log(`${scene.name}: recording`);
		const viewport = scene.viewport ?? { width: 1280, height: 800 };
		const context = await browser.newContext({
			viewport,
			deviceScaleFactor: 1,
			colorScheme: scene.colorScheme ?? 'light',
			recordVideo: { dir: join(work, scene.name), size: viewport },
		});
		const page = await context.newPage();
		const opened = Date.now();
		const errors = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.addInitScript(cursorInit);
		if (scene.init) await page.addInitScript(scene.init);
		let trim = 0;
		let duration = 0;
		page.setDefaultTimeout(30_000);
		let timer;
		const watchdog = new Promise((_, reject) => {
			timer = setTimeout(() => reject(new Error('the scene took longer than 120 s')), 120_000);
		});
		try {
			await Promise.race([
				(async () => {
					await scene.setup(page, server.base);
					await page.evaluate(() => document.fonts.ready);
					await page.waitForTimeout(300);
					trim = (Date.now() - opened) / 1000;
					await scene.play(page);
					duration = (Date.now() - opened) / 1000 - trim;
				})(),
				watchdog,
			]);
		} catch (error) {
			failed = true;
			duration = 0;
			console.error(`${scene.name}: ${error.message}`);
			const shot = join(work, `${scene.name}-failure.png`);
			await page.screenshot({ path: shot, timeout: 5_000 }).then(
				() => console.error(`  screenshot: ${shot}`),
				() => {},
			);
		} finally {
			clearTimeout(timer);
		}
		const video = page.video();
		await context.close();
		if (!duration) continue;
		if (errors.length) console.warn(`${scene.name}: page errors: ${errors.join(' | ')}`);
		const out = join(ROOT, scene.out);
		toGif(await video.path(), out, {
			trim,
			duration,
			width: scene.width ?? 880,
			fps: scene.fps ?? 10,
		});
		const size = statSync(out).size;
		console.log(
			`${relative(ROOT, out)}: ${duration.toFixed(1)} s, ${(size / 1024 / 1024).toFixed(2)} MB`,
		);
		if (size > 2.5 * 1024 * 1024) console.warn(`  over the 2.5 MB budget`);
	}
} finally {
	await browser.close();
	server.stop();
	if (keep) console.log(`Recordings kept in ${work}`);
	else rmSync(work, { recursive: true, force: true });
}
if (failed) process.exit(1);
