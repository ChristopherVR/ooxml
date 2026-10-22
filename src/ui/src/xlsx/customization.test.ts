// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { EditorCore } from './editor-core';
import { defineXlsxEditor, type ExcelCustomization, type XlsxEditorElement } from './index';
import { customizationOf, resolveExcelCustomization } from './customization';
import { openContextMenu } from './context-menu/index';

beforeAll(() => defineXlsxEditor());
afterEach(() => document.body.replaceChildren());

function mount(customization?: ExcelCustomization) {
	const element = document.createElement('xlsx-editor') as XlsxEditorElement;
	if (customization) element.customization = customization;
	document.body.append(element);
	element.newWorkbook();
	return element;
}
const core = (element: XlsxEditorElement) => (element as unknown as { core: EditorCore }).core;
type Backstage = HTMLElement & { items: readonly { id: string; label: string; group?: string }[] };
const backstage = (element: XlsxEditorElement) =>
	element.shadowRoot!.querySelector('office-ui-backstage') as Backstage;
const navIds = (element: XlsxEditorElement) => backstage(element).items.map((item) => item.id);

describe('<xlsx-editor> customization', () => {
	it('resolves capabilities into hidden commands and File tab pages', () => {
		const resolved = resolveExcelCustomization({ capabilities: { print: false, copy: false } });
		expect([...resolved.commands]).toEqual([
			'file.print',
			'edit.copy',
			'edit.cut',
			'home.copy',
			'home.cut',
		]);
		expect([...resolved.backstagePages]).toEqual(['print']);
		expect(resolved.canEdit).toBe(true);
	});

	it('hides and refuses the commands of disabled capabilities, leaving hiddenActions alone', async () => {
		const element = mount({ capabilities: { print: false, find: false } });
		const commands = core(element).commands;
		expect(commands.isHidden('file.print')).toBe(true);
		expect(commands.isHidden('home.find')).toBe(true);
		expect(commands.isHidden('home.bold')).toBe(false);
		expect(element.hiddenActions).toEqual([]);
		expect(await commands.run('file.print')).toBe(false);
		element.customization = null;
		expect(commands.isHidden('file.print')).toBe(false);
	});

	it('edit: false makes the workbook read-only until the capability is removed', () => {
		const element = mount({ capabilities: { edit: false } });
		expect(element.readOnly).toBe(true);
		element.readOnly = false;
		expect(element.readOnly).toBe(true);
		expect(element.hasAttribute('read-only')).toBe(false);
		element.customization = {};
		expect(element.readOnly).toBe(false);
	});

	it('removes File tab pages and adds host pages that mount on show', () => {
		const destroy = vi.fn();
		const element = mount({
			capabilities: { open: false, export: false, settings: false },
			backstage: {
				extraPages: [
					{
						id: 'policies',
						label: 'Policies',
						render: (container, context) => {
							container.textContent = `Policies for ${context.fileName}`;
							return destroy;
						},
					},
				],
			},
		});
		expect(navIds(element)).toEqual(['info', 'save', 'print', 'share', 'host.policies']);
		backstage(element).dispatchEvent(
			new CustomEvent('office-backstage-select', {
				detail: { id: 'host.policies' },
				cancelable: true,
			}),
		);
		const page = backstage(element).querySelector('[data-backstage-page="host.policies"]')!;
		expect(page.textContent).toBe(`Policies for ${element.fileName}`);
		element.customization = null;
		expect(destroy).toHaveBeenCalledTimes(1);
		expect(navIds(element)).toContain('options');
		expect(navIds(element)).not.toContain('host.policies');
	});

	it('gives the cell menu the host commands, which run with the selection', () => {
		const onSelect = vi.fn();
		const element = mount({
			contextMenu: { extraCommands: [{ id: 'send', label: 'Send range', onSelect }] },
		});
		const { ctx } = core(element);
		const commands = customizationOf(ctx).menuCommands;
		expect(commands.map((command) => command.id)).toEqual(['send']);
		const context = { sheet: 0, sheetName: 'Sheet1', ref: 'A1', active: 'A1', readOnly: false };
		const menu = openContextMenu(ctx, [], 0, 0, { host: { commands, context } });
		const { state } = menu.element as HTMLElement & { state: { items: { id: string }[] } };
		expect(state.items.map((item) => item.id)).toEqual(['host:send']);
		menu.element.dispatchEvent(
			new CustomEvent('office-menu-request', { detail: { id: 'host:send' } }),
		);
		expect(onSelect).toHaveBeenCalledWith(context);
		expect(menu.element.isConnected).toBe(false);
	});
});
