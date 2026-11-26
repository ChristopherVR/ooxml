import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { buildSync } from 'esbuild';

export const paneOrigin = 'https://publisher.office-add-in.test';
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
	.replace('DefaultMinVersion="1.10"', 'DefaultMinVersion="1.1"')
	.replace('Name="AddinCommands"', 'Name="TextCoercion"')
	.replace(
		'</Hosts>',
		'</Hosts><Requirements><Sets DefaultMinVersion="1.1"><Set Name="TextCoercion" /></Sets></Requirements>',
	);
export const jsonManifest = fixture('manifest.json')
	.replaceAll('https://notes.example.test', paneOrigin)
	.replace(
		'"name": "AddinCommands", "minVersion": "1.10"',
		'"name": "TextCoercion", "minVersion": "1.1"',
	);

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
