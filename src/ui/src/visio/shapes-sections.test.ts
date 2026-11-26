import { afterEach, describe, expect, it, vi } from 'vitest';
import type { VisioEdit } from 'ooxml-core/visio';
import { demoDocument } from 'ooxml-core/visio/ui';
import { ViewerController } from './controller';
import { createShapesWindow, setShapesLocale } from './shapes-window';
import { SHAPES_STORAGE_KEY } from './shapes-sections';
import { masterCreation, outlinePreviewPath } from './stencil-catalog';
import { wireStencil } from './viewer-stencil';
import type { CancellableEditor } from './worker-editor';

afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

const command = (menu: Element, id: string) =>
	menu.dispatchEvent(new CustomEvent('office-command', { detail: { command: id }, bubbles: true }));
function pane() {
	const element = createShapesWindow(document);
	document.body.append(element);
	const more = element.querySelector('[data-shapes-menu="more"]')!;
	const sections = () =>
		[...element.querySelectorAll<HTMLElement>('#shapes-sections [data-stencil]')].map(
			(section) => section.dataset.stencil,
		);
	const quick = () =>
		[...element.querySelectorAll<HTMLElement>('#shapes-quick [data-master]')].map(
			(button) => button.dataset.master,
		);
	return { element, more, sections, quick };
}

describe('More Shapes and Quick Shapes', () => {
	it('localizes the close label of a restored stencil before opening any menu', () => {
		localStorage.setItem(SHAPES_STORAGE_KEY, JSON.stringify({ open: ['basic-flowchart'] }));
		const { element } = pane();
		const closeLabel = () =>
			element.querySelector('[data-close-stencil="basic-flowchart"]')?.getAttribute('aria-label');
		expect(closeLabel()).toBe('Close Basic Flowchart Shapes');
		setShapesLocale(element, 'de-DE');
		expect(closeLabel()).toBe('Standardflussdiagramm-Shapes schließen');
		expect(element.querySelector('[data-stencil="basic"]')?.getAttribute('aria-label')).toBe(
			'Standard-Shapes',
		);
		expect(element.querySelector('[command="stencil:arrow-shapes"]')?.getAttribute('label')).toBe(
			'Pfeil-Shapes',
		);
		setShapesLocale(element, 'fr');
		expect(closeLabel()).toBe('Fermer Formes de diagramme de flux simple');
		setShapesLocale(element, 'en');
		expect(closeLabel()).toBe('Close Basic Flowchart Shapes');
		expect(element.querySelector('[data-shapes-menu="stencil"]')).toBeNull();
	});
	it('opens a built-in stencil as a collapsible section that can be closed', () => {
		const { element, more, sections } = pane();
		expect(sections()).toEqual(['basic']);
		const items = [...more.querySelectorAll('office-ui-menu-item')];
		expect(items.map((item) => item.getAttribute('label'))).toEqual([
			'Basic Flowchart Shapes',
			'Miscellaneous Flowchart Shapes',
			'Arrow Shapes',
			'New Stencil',
			'Open Stencil...',
		]);
		expect(items[4]!.hasAttribute('disabled')).toBe(false);
		command(more, 'stencil:basic-flowchart');
		expect(sections()).toEqual(['basic', 'basic-flowchart']);
		expect(more.querySelector('[command="stencil:basic-flowchart"]')!.getAttribute('checked')).toBe(
			'true',
		);
		const flowchart = element.querySelector<HTMLElement>('[data-stencil="basic-flowchart"]')!;
		expect(flowchart.querySelector('.stencil-title')!.textContent).toBe('Basic Flowchart Shapes');
		for (const name of [
			'Process',
			'Decision',
			'Subprocess',
			'Start/End',
			'Database',
			'Custom 4',
			'On-page reference',
		])
			expect(flowchart.querySelector(`li[data-name="${name}"] [data-master]`)).not.toBeNull();
		// Collapse and expand, as Visio's stencil title bars do.
		const title = flowchart.querySelector<HTMLButtonElement>('.stencil-title')!;
		title.click();
		expect(title.getAttribute('aria-expanded')).toBe('false');
		expect(flowchart.querySelector<HTMLElement>('.masters')!.hidden).toBe(true);
		title.click();
		expect(flowchart.querySelector<HTMLElement>('.masters')!.hidden).toBe(false);
		// The open stencil persists for this viewer.
		expect(JSON.parse(localStorage.getItem(SHAPES_STORAGE_KEY)!).open).toEqual(['basic-flowchart']);
		expect(pane().sections()).toEqual(['basic', 'basic-flowchart']);
		const close = flowchart.querySelector<HTMLButtonElement>('[data-close-stencil]')!;
		close.focus();
		close.click();
		expect(sections()).toEqual(['basic']);
		expect(document.activeElement).toBe(element.querySelector('#shapes-more'));
	});

	it('shows Quick Shapes and loads the master menu on its first use', async () => {
		const { element, more, quick } = pane();
		const toggle = element.querySelector<HTMLButtonElement>('#shapes-quick-toggle')!;
		expect(quick()).toEqual([]);
		toggle.click();
		expect(toggle.getAttribute('aria-expanded')).toBe('true');
		expect(quick()).toEqual(['rectangle', 'square', 'ellipse', 'circle']);
		command(more, 'stencil:arrow-shapes');
		expect(quick()).toEqual([
			'rectangle',
			'square',
			'ellipse',
			'circle',
			'arrow-right',
			'arrow-left',
			'arrow-up',
			'arrow-down',
		]);
		const masterMenu = element.querySelector('[data-shapes-menu="master"]')!;
		expect(masterMenu.children).toHaveLength(0);
		expect(element.querySelector('[data-shapes-menu="stencil"]')).toBeNull();
		const rightClick = (id: string) =>
			element
				.querySelector(`#shapes-sections [data-master="${id}"]`)!
				.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
		rightClick('star');
		await vi.waitFor(() =>
			expect(masterMenu.querySelector('office-ui-menu-item')?.getAttribute('label')).toBe(
				'Add to Quick Shapes',
			),
		);
		command(masterMenu, 'quick-shapes');
		rightClick('square');
		expect(masterMenu.querySelector('office-ui-menu-item')!.getAttribute('label')).toBe(
			'Remove from Quick Shapes',
		);
		command(masterMenu, 'quick-shapes');
		expect(quick().slice(0, 4)).toEqual(['rectangle', 'ellipse', 'circle', 'star']);
		expect(JSON.parse(localStorage.getItem(SHAPES_STORAGE_KEY)!).quick.basic).toEqual([
			'rectangle',
			'ellipse',
			'circle',
			'star',
		]);
	});

	it('ignores corrupt storage and unknown stencils', () => {
		localStorage.setItem(SHAPES_STORAGE_KEY, '{not json');
		expect(pane().sections()).toEqual(['basic']);
		localStorage.setItem(SHAPES_STORAGE_KEY, JSON.stringify({ open: ['nope', 'arrow-shapes'] }));
		expect(pane().sections()).toEqual(['basic', 'arrow-shapes']);
	});

	it('drops stencil masters as instances with the default Quick Style', async () => {
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
		await controller.load(new Uint8Array([1]));
		const { element, more } = pane();
		command(more, 'stencil:basic-flowchart');
		const viewport = document.createElement('div');
		document.body.append(viewport);
		const messages: string[] = [];
		wireStencil(element, viewport, controller, (message) => messages.push(message));
		const settle = () => new Promise((resolve) => setTimeout(resolve, 0));
		element.querySelector<HTMLButtonElement>('[data-master="flowchart-decision"]')!.click();
		await settle();
		element
			.querySelector<HTMLButtonElement>('[data-master="flowchart-on-page-reference"]')!
			.click();
		await settle();
		// Each drop is an instance of the stencil master, which the core adds to the drawing.
		expect(
			edits.map(([created]) => [created!.type, (created as { master?: string }).master]),
		).toEqual([
			['drop-stencil-master', 'flowchart-decision'],
			['drop-stencil-master', 'flowchart-on-page-reference'],
		]);
		expect(edits.every((batch) => batch[1]?.type === 'format-shape')).toBe(true);
		expect(messages.at(-1)).toMatch(/^On-page reference .+ added from Basic Flowchart Shapes\.$/);
	});

	it('resolves every master and draws previews inside the tile', () => {
		expect(masterCreation('arrow-curved')).toEqual({
			size: { width: 1, height: 0.75 },
			create: { kind: 'rectangle', shape: 'arrow-curved' },
		});
		expect(masterCreation('circle')?.create).toEqual({ kind: 'ellipse' });
		expect(masterCreation('missing')).toBeUndefined();
		const path = outlinePreviewPath(
			{
				paths: [
					[
						[0, 0],
						[1, 0],
						[1, 1],
					],
				],
			},
			{ width: 2, height: 1 },
		);
		expect(path).toBe('M2 17L22 17L22 7Z');
	});
});
