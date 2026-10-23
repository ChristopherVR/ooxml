/**
 * The viewer stylesheet's Tailwind preflight is scoped: it matches only inside
 * `[data-pptx-viewer]`, `[data-pptx-viewport]` and `[data-pptx-scope]`, so a
 * host page keeps its own headings, lists and margins. Every surface this
 * binding teleports to `document.body` leaves the `.pptx-vue-viewer` root, so it
 * must carry `data-pptx-scope` itself or it loses the reset (and renders with
 * browser-default borders, margins and list bullets). A shared control such as
 * `pptx-ui-context-menu` is wrapped in a marked `display: contents` element
 * instead of being marked itself, so its own `:host` typography still applies.
 */
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { mount } from '@vue/test-utils';
import { afterEach, describe, expect, it } from 'vitest';

import {
	findScopedCustomElements,
	scopedCustomElementsIn,
} from '../../../../../scripts/scope-marker-tags.mjs';
import ContextMenu from './ContextMenu.vue';
import ExportProgressModal from './ExportProgressModal.vue';
import ModalDialog from './ModalDialog.vue';

const viewerDir = path.resolve(import.meta.dirname, '..');

function vueFiles(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const file = path.join(dir, entry.name);
		if (entry.isDirectory()) {
			return vueFiles(file);
		}
		return entry.name.endsWith('.vue') ? [file] : [];
	});
}

/**
 * The opening tags of the elements a `<Teleport to="body">` block renders at its
 * top level (`<template>` wrappers are transparent).
 */
function teleportedRoots(source: string): string[] {
	const roots: string[] = [];
	for (const block of source.matchAll(/<Teleport\s+to="body"[^>]*>([\s\S]*?)<\/Teleport>/gu)) {
		let depth = 0;
		// Attribute values may hold `>` (`v-if="rows.length > 0"`), so quoted values are skipped whole.
		const tags = /<(\/?)([a-zA-Z][\w-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/gu;
		for (const tag of (block[1] ?? '').matchAll(tags)) {
			const [whole, closing, name, , selfClosing] = tag;
			if (name === 'template') {
				continue;
			}
			if (closing) {
				depth--;
				continue;
			}
			if (depth === 0) {
				roots.push(whole);
			}
			if (!selfClosing) {
				depth++;
			}
		}
	}
	return roots;
}

afterEach(() => {
	document.body.innerHTML = '';
});

describe('teleported viewer surfaces keep the scoped reset', () => {
	it('marks the root of every surface teleported to document.body', () => {
		const roots = vueFiles(viewerDir).flatMap((file) =>
			teleportedRoots(readFileSync(file, 'utf8')).map((tag) => ({
				file: path.relative(viewerDir, file),
				tag,
			})),
		);
		// ThemeGallery (backdrop and dialog), PresentationMode, ModalDialog, ContextMenu and
		// SettingsDialog.
		expect(roots.length).toBeGreaterThanOrEqual(6);
		for (const { file, tag } of roots) {
			expect(`${file}: ${tag}`).toContain('data-pptx-scope');
		}
	});

	it('replaces the global Tailwind preflight with the viewer-scoped one', () => {
		const css = readFileSync(path.join(viewerDir, '../styles/pptx-vue-viewer.css'), 'utf8');
		expect(css).not.toMatch(/@import\s+'tailwindcss'\s*;/u);
		expect(css).toContain("@import 'tailwindcss/theme.css' layer(theme);");
		expect(css).toContain('src/ui/src/pptx/theme/viewer-preflight.css');
		expect(css).toContain("@import 'tailwindcss/utilities.css' layer(utilities);");
	});

	it('marks the viewer root itself', () => {
		const source = readFileSync(path.join(viewerDir, 'PowerPointViewer.vue'), 'utf8');
		const root = source.match(
			/<div\b(?:[^>"']|"[^"]*"|'[^']*')*\bclass="pptx-vue-viewer"(?:[^>"']|"[^"]*"|'[^']*')*>/u,
		);
		expect(root).not.toBeNull();
		// Each attribute on its own: the root carries others (the theme root) in any order.
		expect(root?.[0]).toMatch(/\sdata-pptx-scope[\s>]/u);
		expect(root?.[0]).toMatch(/\sdata-pptx-theme-root[\s>]/u);
	});

	it('never puts a scope marker on a custom element', () => {
		// An outermost scope takes the page's base text, which beats a shadow-DOM control's own
		// `:host` typography: the marker goes on an ordinary wrapper around the control.
		expect(scopedCustomElementsIn(path.resolve(viewerDir, '..'), ['.vue', '.ts'])).toStrictEqual(
			[],
		);
		expect(
			findScopedCustomElements('<pptx-ui-context-menu v-if="a > b" data-pptx-scope />'),
		).toHaveLength(1);
		expect(findScopedCustomElements('<pptx-ui-x :data-pptx-scope="on"></pptx-ui-x>')).toHaveLength(
			1,
		);
		expect(
			findScopedCustomElements('<div class="contents" data-pptx-scope><pptx-ui-x /></div>'),
		).toStrictEqual([]);
	});

	it('teleports the context menu inside a scoped wrapper, not as a scope itself', () => {
		mount(ContextMenu, {
			props: { open: true, x: 10, y: 10, items: [{ id: 'copy', label: 'Copy' }] },
			attachTo: document.body,
		});
		const menu = document.body.querySelector('pptx-ui-context-menu');
		expect(menu?.hasAttribute('data-pptx-scope')).toBeFalsy();
		expect(menu?.parentElement?.hasAttribute('data-pptx-scope')).toBeTruthy();
		expect(menu?.parentElement?.parentElement).toBe(document.body);
	});

	it('renders a teleported modal inside a scope', () => {
		mount(ModalDialog, { props: { open: true, title: 'Scoped' }, attachTo: document.body });
		const backdrop = document.body.querySelector('.pptx-vue-modal-backdrop');
		expect(backdrop?.parentElement).toBe(document.body);
		expect(backdrop?.hasAttribute('data-pptx-scope')).toBeTruthy();
	});

	it('renders the export progress modal in place, inside the viewer scope', () => {
		// The shared modal is not teleported: it stays under the viewer root, which carries the
		// scope (and the theme), and as a custom element it carries no marker of its own.
		const viewer = document.createElement('div');
		viewer.setAttribute('data-pptx-scope', '');
		document.body.append(viewer);
		mount(ExportProgressModal, {
			props: { open: true, title: 'Exporting', progress: 0.5 },
			attachTo: viewer,
		});
		const modal = document.body.querySelector('pptx-ui-export-progress');
		expect(modal?.closest('[data-pptx-scope]')).toBe(viewer);
		expect(modal?.hasAttribute('data-pptx-scope')).toBeFalsy();
	});
});
