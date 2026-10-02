// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { createEditSession, createWorkbook, saveXlsx } from '@christophervr/xlsx-core';
import type { EditorContext, GridController } from './context';

// The feature modules have their own tests; here they are doubles so the element is tested alone.
const mounted: string[] = [];
vi.mock('./modules', () => ({
	installCommands: () => {
		mounted.push('commands');
		return () => undefined;
	},
	mountGrid: (ctx: EditorContext, container: HTMLElement) => {
		mounted.push('grid');
		const surface = container.ownerDocument.createElement('div');
		surface.tabIndex = 0;
		container.append(surface);
		let zoom = 100;
		const grid: GridController = {
			focus: () => surface.focus(),
			scrollTo: () => undefined,
			invalidate: () => undefined,
			beginEdit: () => undefined,
			commitEdit: () => true,
			cancelEdit: () => undefined,
			isEditing: () => false,
			measureText: (text) => text.length * 7,
			zoom: () => zoom,
			setZoom: (value) => {
				zoom = value;
			},
		};
		ctx.attachGrid(grid);
		return () => undefined;
	},
	mountFormulaBar: () => () => undefined,
	mountSheetTabs: () => () => undefined,
}));

const { XlsxEditorElement, defineXlsxEditor } = await import('./index');
const flush = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

function editor() {
	const element = document.createElement('xlsx-editor') as InstanceType<typeof XlsxEditorElement>;
	const events: { type: string; detail: unknown }[] = [];
	for (const type of [
		'workbook-change',
		'workbook-error',
		'workbook-warning',
		'readonly-change',
		'ribbon-action',
		'file-command',
		'selection-change',
		'sheet-change',
		'dirty-change',
	])
		element.addEventListener(type, (event) =>
			events.push({ type, detail: (event as CustomEvent).detail }),
		);
	document.body.append(element);
	return { element, events, of: (type: string) => events.filter((event) => event.type === type) };
}

async function xlsxBytes(sheets = ['Data', 'Second']) {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.renameSheet(0, sheets[0]!);
	for (const name of sheets.slice(1)) session.addSheet(name);
	session.setCellInput(0, 0, 0, '5');
	return saveXlsx(workbook);
}

beforeAll(() => defineXlsxEditor());

describe('<xlsx-editor> shell', () => {
	it('builds the parts in the shadow root and mounts the modules', () => {
		const { element } = editor();
		const root = element.shadowRoot!;
		for (const part of [
			'title-bar',
			'ribbon',
			'grid',
			'formula-bar',
			'sheet-tabs',
			'status-bar',
			'backstage',
		])
			expect(root.querySelector(`[part~="${part}"]`), part).not.toBeNull();
		expect(mounted).toEqual(expect.arrayContaining(['commands', 'grid']));
		expect(element.getAttribute('aria-label')).toBe('Spreadsheet editor');
	});

	it('loads .xlsx bytes with a session, a clean state and workbook-change', async () => {
		const { element, of } = editor();
		await element.load(await xlsxBytes(), 'Report.xlsx');
		expect(element.workbook?.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Second']);
		expect(element.fileName).toBe('Report.xlsx');
		expect(element.dirty).toBe(false);
		expect(of('workbook-change')).toHaveLength(1);
		expect(element.shadowRoot!.querySelector('.xve-filename')!.textContent).toBe('Report.xlsx');
	});

	it('loads CSV and legacy .xls files, telling the user that Save writes .xlsx', async () => {
		const { element } = editor();
		await element.load(new TextEncoder().encode('a,b\n1,2\n'), 'data.csv');
		expect(element.workbook?.format).toBe('csv');
		const xls = readFileSync(resolve(process.cwd(), 'tests/support/legacy-97.xls'));
		await element.load(new Uint8Array(xls), 'legacy-97.xls');
		expect(element.workbook?.format).toBe('xls');
		await flush();
		expect(element.shadowRoot!.querySelector('.xve-toast')!.textContent).toContain('.xls');
	});

	it('reports load failures through workbook-error and rejects', async () => {
		const { element, of } = editor();
		await expect(
			element.load(new Uint8Array([0x50, 0x4b, 3, 4, 9, 9, 9, 9]), 'broken.xlsx'),
		).rejects.toThrow();
		expect(of('workbook-error')).toHaveLength(1);
	});

	it('creates, edits, saves and marks clean', async () => {
		const { element, of } = editor();
		element.newWorkbook();
		expect(element.workbook?.sheets).toHaveLength(1);
		expect(element.fileName).toBe('Book1.xlsx');
		(
			element as unknown as { core: { session: { setCellInput(...a: unknown[]): void } } }
		).core.session.setCellInput(0, 0, 0, '7');
		expect(element.dirty).toBe(true);
		expect(of('dirty-change').at(-1)?.detail).toEqual({ dirty: true });
		const bytes = await element.saveBytes();
		expect(bytes[0]).toBe(0x50);
		expect(new TextDecoder().decode(await element.saveBytes('csv'))).toContain('7');
		expect((await element.save()).type).toContain('spreadsheetml');
		expect(element.dirty).toBe(true);
		element.markClean();
		expect(element.dirty).toBe(false);
	});

	it('selects ranges, reports them and switches sheets', async () => {
		const { element, of } = editor();
		await element.load(await xlsxBytes(), 'a.xlsx');
		element.setActiveSheet(0);
		element.select('B2:C4');
		expect(element.getSelection()).toBe('B2:C4');
		expect(of('selection-change').at(-1)?.detail).toEqual({ sheet: 0, ref: 'B2:C4', active: 'B2' });
		element.select('Second!A1');
		expect(element.activeSheet).toBe(1);
		expect(of('sheet-change').at(-1)?.detail).toEqual({ index: 1, name: 'Second' });
		expect(() => element.select('not a ref')).toThrow();
		element.setActiveSheet(5);
		expect(element.activeSheet).toBe(1);
	});

	it('undoes and redoes through the edit commands', async () => {
		const { element } = editor();
		element.newWorkbook();
		const session = (
			element as unknown as { core: { session: { setCellInput(...a: unknown[]): void } } }
		).core.session;
		session.setCellInput(0, 0, 0, '1');
		element.undo();
		await flush();
		expect(element.workbook!.sheets[0]!.rows.get(0)?.get(0)).toBeUndefined();
		element.redo();
		await flush();
		expect(element.workbook!.sheets[0]!.rows.get(0)?.get(0)?.value).toBe(1);
	});

	it('runs file commands unless the host cancels file-command', async () => {
		const { element, of } = editor();
		element.newWorkbook();
		element.addEventListener('file-command', (event) => event.preventDefault());
		const newWorkbook = element.workbook;
		await flush();
		element
			.shadowRoot!.querySelector<HTMLButtonElement>('.xve-quick-access [aria-label="Save"]')!
			.click();
		await flush();
		expect(of('file-command').at(-1)?.detail).toEqual({ command: 'save' });
		expect(element.workbook).toBe(newWorkbook);
	});

	it('toggles read-only from the title bar and focuses the grid', async () => {
		const { element, of } = editor();
		element.newWorkbook();
		const mode = element.shadowRoot!.querySelector<HTMLSelectElement>('.xve-mode-select')!;
		mode.value = 'viewing';
		mode.dispatchEvent(new Event('change'));
		expect(element.readOnly).toBe(true);
		expect(element.hasAttribute('read-only')).toBe(true);
		expect(of('readonly-change')).toEqual([
			{ type: 'readonly-change', detail: { readOnly: true } },
		]);
		element.focusGrid();
		expect(element.shadowRoot!.activeElement).not.toBeNull();
	});

	it('opens the backstage from the File tab and closes it with Escape', () => {
		const { element } = editor();
		element.newWorkbook();
		const root = element.shadowRoot!;
		root.querySelector<HTMLButtonElement>('.xve-file-tab')!.click();
		const backstage = root.querySelector<HTMLElement>('[part="backstage"]')!;
		expect(backstage.hidden).toBe(false);
		expect(backstage.querySelector('h2')!.textContent).toBe('Info');
		backstage.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(backstage.hidden).toBe(true);
	});
});

describe('attributes and properties', () => {
	it('reflects attributes both ways without ping-pong and keeps defaults', () => {
		const { element } = editor();
		element.setAttribute('locale', 'fr-CA');
		expect(element.locale).toBe('fr');
		expect(element.shadowRoot!.querySelector('.xve-file-tab')!.textContent).toBe('Fichier');
		element.showFormulaBar = false;
		expect(element.getAttribute('show-formula-bar')).toBe('false');
		element.setAttribute('show-toolbar', 'false');
		expect(element.showToolbar).toBe(false);
		element.fileName = '';
		expect(element.fileName).toBe('Book1.xlsx');
		element.authorName = '';
		expect(element.authorName).toBe('Author');
		element.setAttribute('theme', 'sepia');
		expect(element.theme).toBe('auto');
	});

	it('applies themeColors inline and clears stale overrides', () => {
		const { element } = editor();
		element.themeColors = { primary: '#ff0000', gridLine: '#00ff00' };
		expect(element.style.getPropertyValue('--xve-primary')).toBe('#ff0000');
		expect(element.style.getPropertyValue('--xve-grid-line')).toBe('#00ff00');
		element.themeColors = {};
		expect(element.style.getPropertyValue('--xve-primary')).toBe('');
		expect(element.themeColors).toEqual({});
	});

	it('accepts a workbook model without echoing workbook-change', () => {
		const { element, of } = editor();
		const workbook = createWorkbook();
		element.workbook = workbook;
		expect(element.workbook).toBe(workbook);
		expect(element.dirty).toBe(false);
		expect(of('workbook-change')).toHaveLength(0);
		element.workbook = null;
		expect(element.workbook).toBeNull();
	});
});
