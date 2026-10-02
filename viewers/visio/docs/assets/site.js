const root = document.documentElement;
const themeButton = document.querySelector('.theme-toggle');
function syncThemeLabel() {
	if (!themeButton) return;
	const dark =
		root.dataset.theme === 'dark' ||
		(!root.dataset.theme && matchMedia('(prefers-color-scheme: dark)').matches);
	themeButton.setAttribute('aria-label', `Switch to ${dark ? 'light' : 'dark'} theme`);
	themeButton.setAttribute('title', `Switch to ${dark ? 'light' : 'dark'} theme`);
	themeButton.textContent = dark ? '☀' : '◐';
}
try {
	const saved = localStorage.getItem('visio-docs-theme');
	if (saved === 'dark' || saved === 'light') root.dataset.theme = saved;
} catch {
	/* Preference storage can be unavailable. */
}
syncThemeLabel();
themeButton?.addEventListener('click', () => {
	const dark =
		root.dataset.theme === 'dark' ||
		(!root.dataset.theme && matchMedia('(prefers-color-scheme: dark)').matches);
	root.dataset.theme = dark ? 'light' : 'dark';
	try {
		localStorage.setItem('visio-docs-theme', root.dataset.theme);
	} catch {
		/* Keep session preference without storage. */
	}
	syncThemeLabel();
});
const snippets = {
	vanilla: [
		'main.ts',
		`import { mountViewer } from '@christophervr/visio-viewer';\n\nconst viewer = mountViewer(host, {\n  events: {\n    'shape-select': (shape) => console.log(shape),\n  },\n});\n\nawait viewer.load(file); // A local File or bytes\nviewer.fit();\n\n// When your view unmounts:\nviewer.destroy();`,
	],
	react: [
		'Diagram.tsx · lifecycle pattern',
		`import { useEffect, useRef } from 'react';\nimport { mountViewer } from '@christophervr/visio-viewer';\n\nexport function Diagram({ file }: { file: File }) {\n  const host = useRef<HTMLDivElement>(null);\n  useEffect(() => {\n    if (!host.current) return;\n    const viewer = mountViewer(host.current);\n    viewer.load(file).catch(console.error);\n    return () => viewer.destroy();\n  }, [file]);\n  return <div ref={host} style={{ height: 600 }} />;\n}`,
	],
	vue: [
		'Diagram.vue · lifecycle pattern',
		`<script setup lang="ts">\nimport { ref, onMounted, onBeforeUnmount } from 'vue';\nimport { mountViewer } from '@christophervr/visio-viewer';\nconst props = defineProps<{ file: File }>();\nconst host = ref<HTMLDivElement>();\nlet viewer: ReturnType<typeof mountViewer>;\nonMounted(() => {\n  viewer = mountViewer(host.value!);\n  viewer.load(props.file).catch(console.error);\n});\nonBeforeUnmount(() => viewer?.destroy());\n</script>\n<template><div ref="host" style="height:600px" /></template>`,
	],
	angular: [
		'Component lifecycle · integration pattern',
		`import { mountViewer } from '@christophervr/visio-viewer';\n\n// Inside your component, after the host view exists:\nngAfterViewInit() {\n  this.viewer = mountViewer(this.host.nativeElement);\n  this.viewer.load(this.file).catch(console.error);\n}\n\n// Release subscriptions and pending work:\nngOnDestroy() {\n  this.viewer?.destroy();\n}\n\n// host: ElementRef<HTMLElement>; file: File`,
	],
	svelte: [
		'Diagram.svelte · lifecycle pattern',
		`<script lang="ts">\nimport { onMount } from 'svelte';\nimport { mountViewer } from '@christophervr/visio-viewer';\nlet { file }: { file: File } = $props();\nlet host: HTMLDivElement;\nonMount(() => {\n  const viewer = mountViewer(host);\n  viewer.load(file).catch(console.error);\n  return () => viewer.destroy();\n});\n</script>\n<div bind:this={host} style="height:600px"></div>`,
	],
};
const code = document.querySelector('#integration-code');
const filename = document.querySelector('#integration-filename');
const status = document.querySelector('#copy-status');
for (const button of document.querySelectorAll('[data-framework]')) {
	button.addEventListener('click', () => {
		const sample = snippets[button.dataset.framework];
		if (!sample || !code || !filename) return;
		for (const choice of document.querySelectorAll('[data-framework]'))
			choice.setAttribute('aria-pressed', String(choice === button));
		filename.textContent = sample[0];
		code.textContent = sample[1];
		if (status) status.textContent = '';
	});
}
document.querySelector('.copy-code')?.addEventListener('click', async () => {
	try {
		await navigator.clipboard.writeText(code?.textContent ?? '');
		if (status) status.textContent = 'Code copied.';
	} catch {
		if (status) status.textContent = 'Copy is unavailable. Select and copy the code below.';
	}
});
