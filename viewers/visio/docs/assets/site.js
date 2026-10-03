import { initVisioTheme } from './theme.js';
import './search.js';

document.documentElement.classList.add('js');
initVisioTheme(window);

const outline = document.querySelector('.doc-outline nav');
if (outline) {
	for (const [index, heading] of [...document.querySelectorAll('.doc-content h2')].entries()) {
		if (!heading.id) heading.id = `section-${index + 1}`;
		const link = document.createElement('a');
		link.href = `#${heading.id}`;
		link.textContent = heading.textContent;
		outline.append(link);
	}
}

const compactDocs = window.matchMedia('(max-width: 959px)');
const docsMenus = [...document.querySelectorAll('.doc-mobile-menu, .doc-mobile-outline')];
function sizeDocsMenus() {
	for (const menu of docsMenus) menu.open = !compactDocs.matches;
}
sizeDocsMenus();
compactDocs.addEventListener?.('change', sizeDocsMenus);
for (const menu of docsMenus) {
	menu.addEventListener('toggle', () => {
		if (compactDocs.matches && menu.open)
			for (const other of docsMenus) if (other !== menu) other.open = false;
	});
	menu.addEventListener('click', (event) => {
		if (compactDocs.matches && event.target.closest('a')) menu.open = false;
	});
}
document.addEventListener('keydown', (event) => {
	if (event.key === 'Escape' && compactDocs.matches) {
		for (const menu of docsMenus) {
			if (menu.open) {
				menu.open = false;
				menu.querySelector('summary').focus();
			}
		}
	}
});

const menuButton = document.querySelector('.menu-toggle');
const navigation = document.getElementById('site-navigation');
function setMenu(open, returnFocus = false) {
	menuButton?.setAttribute('aria-expanded', String(open));
	if (navigation) navigation.dataset.open = String(open);
	if (returnFocus) menuButton?.focus();
}
menuButton?.addEventListener('click', () =>
	setMenu(menuButton.getAttribute('aria-expanded') !== 'true'),
);
navigation?.addEventListener('click', (event) => {
	if (event.target.closest('a')) setMenu(false);
});
document.addEventListener('keydown', (event) => {
	if (event.key === 'Escape')
		for (const resources of document.querySelectorAll('.nav-resources')) resources.open = false;
	if (event.key === 'Escape' && menuButton?.getAttribute('aria-expanded') === 'true')
		setMenu(false, true);
});

const snippets = {
	vanilla: [
		'main.ts',
		`import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
  events: { 'document-error': console.error },
});

await viewer.load(file); // A local File or bytes
viewer.fit();

// When your view unmounts:
viewer.destroy();`,
	],
	react: [
		'Diagram.tsx · React component',
		`import { VisioViewer } from 'visio-react-viewer';
import type { VisioDocument } from 'visio-react-viewer';

export function Diagram({ diagram }: {
  diagram: VisioDocument;
}) {
  return <VisioViewer
    document={diagram}
    style={{ height: 600 }}
    events={{ 'document-error': console.error }}
  />;
}
// The native component handles mount and cleanup.`,
	],
	vue: [
		'Diagram.vue · Vue component',
		`<script setup lang="ts">
import { VisioViewer } from 'visio-vue-viewer';
import type { VisioDocument } from 'visio-vue-viewer';
defineProps<{ diagram: VisioDocument }>();
</script>

<template>
  <VisioViewer
    :document="diagram"
    style="height: 600px"
    @document-error="console.error"
  />
</template>`,
	],
	angular: [
		'diagram.ts · Angular component',
		`import { Component, Input } from '@angular/core';
import { VisioViewerComponent }
  from 'visio-angular-viewer';
import type { VisioDocument } from 'visio-angular-viewer';

@Component({
  selector: 'app-diagram',
  imports: [VisioViewerComponent],
  template: '<visio-viewer-host [document]="diagram" />',
  styles: ['visio-viewer-host { display:block; height:600px }'],
})
export class Diagram {
  @Input() diagram: VisioDocument | null = null;
}`,
	],
	svelte: [
		'Diagram.svelte · Svelte component',
		`<script lang="ts">
import VisioViewer
  from 'visio-svelte-viewer';
import type { VisioDocument } from 'visio-svelte-viewer';
let { diagram }: { diagram: VisioDocument } = $props();
</script>

<VisioViewer
  document={diagram}
  style="height:600px"
  events={{ 'document-error': console.error }}
/>`,
	],
	solid: [
		'Diagram.tsx · Solid component',
		`import { VisioViewer } from 'visio-solid-viewer';
import type { VisioDocument } from 'visio-solid-viewer';

export function Diagram(props: { diagram: VisioDocument }) {
  return <VisioViewer
    document={props.diagram}
    style={{ height: '600px' }}
    events={{ 'document-error': console.error }}
  />;
}
// Each binding uses the same shared browser element.`,
	],
};
const code = document.getElementById('integration-code');
const filename = document.getElementById('integration-filename');
const status = document.getElementById('copy-status');
const panel = document.getElementById('integration-example');
const choices = [...document.querySelectorAll('[data-framework]')];
let copyAttempt = 0;
function selectFramework(button, focus = false) {
	const sample = snippets[button.dataset.framework];
	if (!sample || !code || !filename) return;
	copyAttempt++;
	for (const choice of choices) {
		choice.setAttribute('aria-selected', String(choice === button));
		choice.tabIndex = choice === button ? 0 : -1;
	}
	panel?.setAttribute('aria-labelledby', button.id);
	filename.textContent = sample[0];
	code.textContent = sample[1];
	if (status) status.textContent = '';
	if (focus) button.focus();
}
for (const button of choices) {
	button.addEventListener('click', () => selectFramework(button));
	button.addEventListener('keydown', (event) => {
		const index = choices.indexOf(button);
		let next;
		if (event.key === 'ArrowRight') next = choices[(index + 1) % choices.length];
		if (event.key === 'ArrowLeft') next = choices[(index + choices.length - 1) % choices.length];
		if (event.key === 'Home') next = choices[0];
		if (event.key === 'End') next = choices.at(-1);
		if (!next) return;
		event.preventDefault();
		selectFramework(next, true);
	});
}
document.querySelector('.copy-code')?.addEventListener('click', async () => {
	const attempt = ++copyAttempt;
	try {
		await navigator.clipboard.writeText(code?.textContent ?? '');
		if (status && attempt === copyAttempt) status.textContent = 'Code copied.';
	} catch {
		if (status && attempt === copyAttempt)
			status.textContent = 'Copy is unavailable. Select and copy the code below.';
	}
});
