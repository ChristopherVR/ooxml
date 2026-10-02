// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { flush, shellFixture, spyCommand } from './test-support/shell';
import { dialogHost } from './dialogs';

describe('command registry', () => {
	it('runs a registered command with its argument and announces ribbon-action', async () => {
		const { core, events } = shellFixture();
		const bold = spyCommand('home.bold');
		core.commands.register(bold);
		expect(await core.commands.run('home.bold', 'x')).toBe(true);
		expect(bold.runs).toEqual(['x']);
		expect(events.filter((e) => e.type === 'ribbon-action')).toEqual([
			{ type: 'ribbon-action', detail: { id: 'home.bold' } },
		]);
	});

	it('refuses missing, disabled, hidden and read-only editing commands', async () => {
		const { core } = shellFixture();
		const editing = spyCommand('a', { editing: true });
		const disabled = spyCommand('b', { enabled: () => false });
		const hidden = spyCommand('c');
		core.commands.registerAll([editing, disabled, hidden]);
		core.hiddenActions = ['c'];
		core.setReadOnly(true);
		expect(await core.commands.run('missing')).toBe(false);
		expect(await core.commands.run('a')).toBe(false);
		expect(await core.commands.run('b')).toBe(false);
		expect(await core.commands.run('c')).toBe(false);
		expect([editing.runs, disabled.runs, hidden.runs]).toEqual([[], [], []]);
		core.setReadOnly(false);
		expect(core.commands.isEnabled('a')).toBe(true);
	});

	it('treats a throwing enabled() as disabled and reports a failing run', async () => {
		const { core, events } = shellFixture();
		core.commands.register(
			spyCommand('x', {
				enabled: () => {
					throw new Error('nope');
				},
			}),
		);
		core.commands.register({
			id: 'y',
			label: 'Y',
			run: () => {
				throw new Error('boom');
			},
		});
		expect(core.commands.isEnabled('x')).toBe(false);
		expect(await core.commands.run('y')).toBe(false);
		expect(events.find((e) => e.type === 'workbook-error')?.detail).toMatchObject({
			message: 'boom',
		});
	});

	it('a later registration replaces an earlier one with the same id', async () => {
		const { core } = shellFixture();
		const first = spyCommand('edit.undo');
		const second = spyCommand('edit.undo');
		core.commands.register(first);
		core.commands.register(second);
		await core.commands.run('edit.undo');
		expect([first.runs.length, second.runs.length]).toEqual([0, 1]);
		expect(core.commands.list().filter((c) => c.id === 'edit.undo')).toHaveLength(1);
	});
});

describe('dialog registry', () => {
	it('opens by name, resolves undefined for unknown, cancelled or failing dialogs', async () => {
		const { core } = shellFixture();
		core.dialogs.register('zoom', async (_ctx, props) => ({ zoom: 150, props }));
		core.dialogs.register('cancel', async () => undefined);
		core.dialogs.register('fail', async () => {
			throw new Error('bad');
		});
		expect(await core.dialogs.open('zoom', 1)).toEqual({ zoom: 150, props: 1 });
		expect(await core.dialogs.open('nope')).toBeUndefined();
		expect(await core.dialogs.open('cancel')).toBeUndefined();
		expect(await core.dialogs.open('fail')).toBeUndefined();
	});

	it('gives dialogs one host element in the shadow root', () => {
		const { core } = shellFixture();
		const a = dialogHost(core.ctx);
		expect(dialogHost(core.ctx)).toBe(a);
		expect(core.ctx.root.contains(a)).toBe(true);
	});
});

describe('editor context', () => {
	it('announces selection changes once and switches sheets with sheet-change', async () => {
		const { core, events } = shellFixture();
		core.session!.addSheet('Second');
		core.ctx.selection.set({ active: { row: 1, col: 1 } });
		core.ctx.emit('selection-change', { sheet: 0, ref: 'B2', active: 'B2' });
		expect(events.filter((e) => e.type === 'selection-change')).toHaveLength(1);
		core.ctx.setActiveSheet(1);
		expect(core.ctx.activeSheet()).toBe(1);
		expect(events.find((e) => e.type === 'sheet-change')?.detail).toEqual({
			index: 1,
			name: 'Second',
		});
		expect(core.ctx.selection.get()).toMatchObject({ sheet: 1, active: { row: 0, col: 0 } });
		await flush();
	});

	it('marks the workbook dirty on edits and fans model changes out', () => {
		const { core, events } = shellFixture();
		const seen: unknown[] = [];
		core.ctx.onModelChange((change) => seen.push(change));
		core.session!.setCellInput(0, 0, 0, '42');
		expect(core.dirty.dirty).toBe(true);
		expect(events.some((e) => e.type === 'dirty-change')).toBe(true);
		expect(events.some((e) => e.type === 'workbook-change')).toBe(true);
		expect(seen).toHaveLength(1);
	});

	it('translates through the locale', () => {
		const { core } = shellFixture();
		core.setLocale('de-DE');
		expect(core.ctx.t('Save')).toBe('Speichern');
		expect(core.ctx.t('{count} compatibility notes', { count: 3 })).toBe(
			'3 Kompatibilitätshinweise',
		);
		expect(core.ctx.t('Unknown key {x}', { x: 1 })).toBe('Unknown key 1');
	});
});
