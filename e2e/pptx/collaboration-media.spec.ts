/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * A participant who joins a shared deck sees its pictures and video, not just its text.
 *
 * The sharer opens a deck with an embedded picture and an embedded video, then a second
 * participant joins the serverless room with nothing but the invitation. The joiner has no
 * copy of the file, so it can only show those parts from the bytes the room carries: each
 * source it renders must be a URL of its own page that yields the same bytes. A deck that
 * links its pictures instead stays a link, and a joiner that blocks external content asks
 * the linked origin for nothing, whatever the sharer allows.
 */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { ImageBuilder, MediaBuilder, PptxHandler } from 'pptx-viewer-core';

import { fixture } from './support/deck';
import { EXTERNAL_ORIGIN, externalContentDeck } from './support/external-content-deck';
import { webrtcRoomPath } from './support/webrtc-room';

/** A 2x2 solid blue PNG. */
const PICTURE = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNkaGBgYGBgYGAAAAN4AE0r1L6tAAAAAElFTkSuQmCC',
	'base64',
);

async function mediaDeck(): Promise<{ name: string; mimeType: string; buffer: Buffer }> {
	const video = await readFile(fixture('media/tiny-video.mp4'));
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	data.slides[0].elements = [
		ImageBuilder.create(`data:image/png;base64,${PICTURE.toString('base64')}`)
			.position(60, 80)
			.size(240, 240)
			.build(),
		MediaBuilder.video(`data:video/mp4;base64,${video.toString('base64')}`)
			.position(420, 80)
			.size(320, 180)
			.build(),
	];
	data.slides[0].isDirty = true;
	const buffer = Buffer.from(await handler.save(data.slides));
	handler.dispose();
	return {
		name: 'shared-media.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer,
	};
}

function viewport(page: Page) {
	return page.locator('[data-pptx-viewport]');
}

async function join(page: Page, roomId: string, name: string, query = ''): Promise<void> {
	await page.goto(webrtcRoomPath(roomId, `${query}name=${name}&`));
	await page.locator('[aria-roledescription="slide"]').first().waitFor({ timeout: 30_000 });
	await expect(
		page.getByRole('status', { name: 'Collaboration: Connected', exact: true }),
	).toBeVisible({ timeout: 30_000 });
}

async function openThroughBackstage(page: Page, file: Awaited<ReturnType<typeof mediaDeck>>) {
	await page
		.getByRole('toolbar', { name: 'Presentation toolbar' })
		.getByRole('tab', { name: 'File', exact: true })
		.click();
	const backstage = page.getByRole('dialog', { name: 'File', exact: true });
	await backstage.getByRole('button', { name: 'Open', exact: true }).click();
	const chooser = page.waitForEvent('filechooser');
	await backstage.getByRole('button', { name: /browse this device/iu }).click();
	await (await chooser).setFiles(file);
}

interface RenderedSource {
	tag: string;
	url: string;
	/** Bytes the URL yields when fetched from its own page, or -1 when it cannot be read. */
	size: number;
}

/** Every picture and media source the slide on screen renders, with the bytes behind it. */
function renderedSources(page: Page): Promise<RenderedSource[]> {
	return viewport(page)
		.first()
		.evaluate(async (root) => {
			const found: { tag: string; url: string }[] = [];
			const visit = (node: Element | ShadowRoot) => {
				for (const element of node.querySelectorAll('img, image, video, source, audio')) {
					const url =
						element.getAttribute('src') ??
						element.getAttribute('href') ??
						element.getAttribute('xlink:href');
					if (url) {
						found.push({ tag: element.tagName.toLowerCase(), url });
					}
				}
				for (const element of node.querySelectorAll('*')) {
					if (element.shadowRoot) {
						visit(element.shadowRoot);
					}
				}
			};
			visit(root);
			return Promise.all(
				found.map(async (source) => {
					try {
						const response = await fetch(source.url);
						return { ...source, size: (await response.arrayBuffer()).byteLength };
					} catch {
						return { ...source, size: -1 };
					}
				}),
			);
		});
}

test.describe('collaboration media', () => {
	test.setTimeout(120_000);

	test('a joiner sees the sharer’s pictures and video from its own URLs', async ({
		page,
	}, info) => {
		const video = await readFile(fixture('media/tiny-video.mp4'));
		const peer = await page.context().newPage();
		const roomId = `e2e-media-${info.project.name}-${Date.now()}`;
		try {
			await join(page, roomId, 'host', 'sample=1&');
			await openThroughBackstage(page, await mediaDeck());
			const sizes = (sources: RenderedSource[]) => sources.map((source) => source.size);
			await expect
				.poll(async () => sizes(await renderedSources(page)), { timeout: 30_000 })
				.toEqual(expect.arrayContaining([PICTURE.length, video.length]));
			const shared = (await renderedSources(page)).map((source) => source.url);

			await join(peer, roomId, 'peer');
			await expect
				.poll(async () => sizes(await renderedSources(peer)), { timeout: 30_000 })
				.toEqual(expect.arrayContaining([PICTURE.length, video.length]));
			const seen = await renderedSources(peer);
			for (const source of seen.filter((item) => item.size > 0)) {
				// A URL of the joiner's own page, never one only the sharer's page can resolve.
				expect(shared).not.toContain(source.url);
				expect(source.url).toMatch(/^(?:blob|data):/u);
			}
			await peer.screenshot({ path: info.outputPath('joiner.png') });
		} finally {
			await peer.close();
		}
	});

	test('a joiner that blocks external content asks the linked origin for nothing', async ({
		page,
	}, info) => {
		const peer = await page.context().newPage();
		const sharerRequests: string[] = [];
		const requests: string[] = [];
		await page.route(`${EXTERNAL_ORIGIN}/**`, (route) => {
			sharerRequests.push(route.request().url());
			return route.fulfill({ status: 404, body: '' });
		});
		await peer.route(`${EXTERNAL_ORIGIN}/**`, (route) => {
			requests.push(route.request().url());
			return route.fulfill({ status: 404, body: '' });
		});
		const roomId = `e2e-linked-${info.project.name}-${Date.now()}`;
		try {
			await join(page, roomId, 'host', 'sample=1&allowExternalContent=1&');
			await openThroughBackstage(page, {
				name: 'linked-content.pptx',
				mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
				buffer: await externalContentDeck(),
			});
			// The sharer allows linked content, so it fetches it.
			await expect.poll(() => sharerRequests.length, { timeout: 30_000 }).toBeGreaterThan(0);

			await join(peer, roomId, 'peer', 'allowExternalContent=0&');
			// The joiner shows the shared slide (the sharer's linked deck has five elements).
			await expect
				.poll(() => viewport(peer).first().locator('[data-element-id]').count(), {
					timeout: 30_000,
				})
				.toBeGreaterThanOrEqual(5);
			await peer.waitForLoadState('networkidle');
			expect(
				await peer.evaluate(
					(origin) => document.documentElement.outerHTML.includes(origin),
					EXTERNAL_ORIGIN,
				),
			).toBe(false);
			expect(requests).toEqual([]);
		} finally {
			await peer.close();
		}
	});
});
