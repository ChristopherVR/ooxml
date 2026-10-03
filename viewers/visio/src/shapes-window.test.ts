import { afterEach, describe, expect, it } from 'vitest';
import type { VisioEdit } from 'ooxml-core/visio';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';
import { BASIC_SHAPES, createShapesWindow } from './shapes-window.js';
import { wireStencil } from './viewer-stencil.js';
import type { CancellableEditor } from './worker-editor.js';

afterEach(() => document.body.replaceChildren());

async function setup(source = true) {
	const edits: VisioEdit[][] = [];
	const editor: CancellableEditor = async (_bytes, commands) => {
		edits.push([...commands]);
		return {
			bytes: new Uint8Array([2]),
			document: structuredClone(demoDocument),
			changedParts: ['visio/pages/page1.xml'],
			diagnostics: [],
		};
	};
	const controller = new ViewerController(
		async () => structuredClone(demoDocument),
		() => {},
		editor,
	);
	if (source) await controller.load(new Uint8Array([1]));
	else controller.setDocument(structuredClone(demoDocument));
	const pane = createShapesWindow(document);
	const viewport = document.createElement('div');
	document.body.append(pane, viewport);
	const messages: string[] = [];
	const dispose = wireStencil(pane, viewport, controller, (message) => messages.push(message));
	const master = (id: string) => pane.querySelector<HTMLButtonElement>(`[data-master="${id}"]`)!;
	const settle = () => new Promise((resolve) => setTimeout(resolve, 0));
	return { controller, pane, edits, messages, master, settle, dispose };
}

describe("Visio's Shapes window", () => {
	it('shows the Basic Shapes stencil with only creatable masters enabled', async () => {
		const { pane, master } = await setup();
		expect(pane.querySelector('.stencil-title')!.textContent).toBe('Basic Shapes');
		expect(pane.querySelectorAll('#shapes-stencils [data-master]')).toHaveLength(
			BASIC_SHAPES.length,
		);
		expect(master('rectangle').draggable).toBe(true);
		expect(master('rectangle').hasAttribute('data-unsupported')).toBe(false);
		expect(master('circle').getAttribute('aria-disabled')).toBe('true');
		expect(master('circle').title).toMatch(/not available yet\. Needs core master drops/);
		expect(
			[...pane.querySelectorAll<HTMLButtonElement>('.shapes-row')].every((row) => row.disabled),
		).toBe(true);
	});

	it('adds a Rectangle at the page centre through core and ignores unsupported masters', async () => {
		const { controller, edits, master, messages, settle } = await setup();
		master('circle').click();
		await settle();
		expect(edits).toEqual([]);
		master('rectangle').click();
		await settle();
		const page = demoDocument.pages[0]!;
		expect(edits).toEqual([
			[
				{
					type: 'create-rectangle',
					pageId: page.id,
					shapeId: expect.any(String),
					x: page.width / 2,
					y: page.height / 2,
					width: 1,
					height: 0.75,
				},
			],
		]);
		expect(messages.at(-1)).toMatch(/^Rectangle .+ added from Basic Shapes\.$/);
		expect(controller.state.edit.canUndo).toBe(true);
	});

	it('explains read-only documents instead of editing them', async () => {
		const { master, edits, messages, settle } = await setup(false);
		master('rectangle').click();
		await settle();
		expect(edits).toEqual([]);
		expect(messages).toEqual([
			'Open a .vsdx file to add shapes. Model-only documents are read only.',
		]);
	});

	it('switches to Search and filters the stencil by name', async () => {
		const { pane } = await setup();
		pane.querySelector<HTMLButtonElement>('[data-shapes-view="search"]')!.click();
		expect(pane.querySelector<HTMLElement>('#shapes-stencils')!.hidden).toBe(true);
		const field = pane.querySelector<HTMLInputElement>('.shapes-search-field')!;
		expect(document.activeElement).toBe(field);
		field.value = 'star';
		field.dispatchEvent(new Event('input'));
		const visible = [...pane.querySelectorAll<HTMLElement>('#shapes-search li')].filter(
			(item) => !item.hidden,
		);
		expect(visible.map((item) => item.dataset.name)).toEqual(['5-point star']);
		field.value = 'zzz';
		field.dispatchEvent(new Event('input'));
		expect(pane.querySelector<HTMLElement>('.shapes-empty')!.hidden).toBe(false);
	});
});
