import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { buildSync } from 'esbuild';
import type { Page } from '@playwright/test';

export const paneOrigin = 'https://publisher.office-add-in.test';
export async function enforcePublisherFramePolicy(page: Page) {
	await page.route('**/*', async (route) => {
		if (!route.request().isNavigationRequest() || route.request().frame() !== page.mainFrame())
			return route.fallback();
		const response = await route.fetch();
		return route.fulfill({
			response,
			headers: {
				...response.headers(),
				'content-security-policy': `frame-src ${paneOrigin}`,
			},
		});
	});
}

export const addInName = 'Notes';
const singlePixelPng = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGPitdoAAAFXAPqIZaQSAAAAAElFTkSuQmCC',
	'base64',
);
const fixture = (name: string) =>
	readFileSync(
		new URL(`../../src/core/automation/add-ins/__fixtures__/${name}`, import.meta.url),
		'utf8',
	);

export const xmlManifest = fixture('manifest.xml')
	.replaceAll('https://notes.example.test', paneOrigin)
	.replace(/<Requirements><bt:Sets .*?<\/Requirements>/, '');
export const jsonManifest = fixture('manifest.json')
	.replaceAll('https://notes.example.test', paneOrigin)
	.replace(/^[\t ]*"requirements":\s*\{\s*"capabilities":.+\r?\n/m, '');

export const packageFiles = [
	['manifest.json', jsonManifest],
	['icons/color.png', singlePixelPng],
	['icons/outline.png', singlePixelPng],
	['fr.json', fixture('fr.json')],
] as const;

export const publisherPage = (editorOrigin: string) => `<!doctype html>
	<title>Local fixture add-in</title>
	<p id="ready" role="status"></p>
	<p id="isolated"></p>
	<button id="read">Read selection</button>
	<button id="write">Write selection</button>
	<output id="result"></output>
	<script src="/client.js" data-editor-origin="${editorOrigin}"></script>`;

let script: string | undefined;
export function publisherScript() {
	script ??= buildSync({
		entryPoints: [fileURLToPath(new URL('./publisher.ts', import.meta.url))],
		bundle: true,
		platform: 'browser',
		format: 'iife',
		write: false,
	}).outputFiles[0]?.text;
	if (!script) throw new Error('Fixture client build did not produce JavaScript.');
	return script;
}
