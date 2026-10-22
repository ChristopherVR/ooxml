import { describe, expect, it, vi } from 'vitest';
import {
	acceptedHostBackstagePages,
	capabilityHiddenActions,
	disabledCapabilities,
	hostBackstageItemId,
	hostBackstagePageId,
	hostBackstageRegionId,
	isCapabilityEnabled,
	isHostMenuEntry,
	mergeHiddenActions,
	placeHostBackstagePages,
	resolveHostMenuCommands,
	sameRegionContext,
	sameRenderer,
	withHostMenuCommands,
	type EditorCapabilities,
	type HostBackstagePage,
	type HostMenuCommand,
} from './index';

describe('capabilities', () => {
	it('treats an omitted or true capability as enabled', () => {
		expect(isCapabilityEnabled(undefined, 'edit')).toBe(true);
		expect(isCapabilityEnabled({ edit: true }, 'edit')).toBe(true);
		const printOff: EditorCapabilities = { print: false };
		expect(isCapabilityEnabled(printOff, 'edit')).toBe(true);
		expect(isCapabilityEnabled({ edit: false }, 'edit')).toBe(false);
	});

	it('lists the disabled ids and the actions they remove', () => {
		const caps = { print: false, export: true, copy: false } as const;
		expect(disabledCapabilities(caps)).toEqual(['print', 'copy']);
		const map = { print: ['print', 'quick-print'], copy: ['copy', 'cut'], export: ['export'] };
		expect(capabilityHiddenActions(caps, map)).toEqual(['print', 'quick-print', 'copy', 'cut']);
		expect(capabilityHiddenActions(undefined, map)).toEqual([]);
	});

	it('merges hidden-action lists without repeats', () => {
		expect(mergeHiddenActions(['a', 'b'], undefined, ['b', 'c'], null)).toEqual(['a', 'b', 'c']);
	});
});

describe('host backstage pages', () => {
	const page = (id: string, extra: Partial<HostBackstagePage<string>> = {}) => ({
		id,
		label: id,
		render: `x-${id}`,
		...extra,
	});

	it('drops empty and repeated ids and sorts by order, stably', () => {
		const pages = [
			page('b', { order: 2 }),
			page(''),
			page('a'),
			page('b', { label: 'second' }),
			page('c', { order: 2 }),
			page('d', { order: -1 }),
		];
		expect(acceptedHostBackstagePages(pages).map((p) => p.id)).toEqual(['d', 'a', 'b', 'c']);
		expect(acceptedHostBackstagePages(pages).find((p) => p.id === 'b')?.label).toBe('b');
	});

	it('places each page after the built-in pages of its group', () => {
		const items = [
			{ id: 'info', group: 'main' as const },
			{ id: 'print' },
			{ id: 'options', group: 'footer' as const },
		];
		const placed = placeHostBackstagePages(
			items,
			[page('about', { group: 'footer' }), page('policies'), page('help')],
			(p) => ({ id: hostBackstageItemId(p.id), ...(p.group ? { group: p.group } : {}) }),
		);
		expect(placed.map((entry) => entry.id)).toEqual([
			'info',
			'print',
			'host.policies',
			'host.help',
			'options',
			'host.about',
		]);
		expect(items).toHaveLength(3);
	});

	it('round-trips nav and region ids', () => {
		expect(hostBackstagePageId(hostBackstageItemId('a.b'))).toBe('a.b');
		expect(hostBackstagePageId('info')).toBeUndefined();
		expect(hostBackstageRegionId('about')).toBe('backstage.host.about');
	});
});

describe('host menu commands', () => {
	const command = (
		id: string,
		extra: Partial<HostMenuCommand<{ n: number }>> = {},
	): HostMenuCommand<{ n: number }> => ({
		id,
		label: id.toUpperCase(),
		onSelect: vi.fn(),
		...extra,
	});

	it('isolates ids, skips repeats and hidden commands, and runs with the context', () => {
		const run = vi.fn();
		const { top, bottom } = resolveHostMenuCommands(
			[
				command('send', { onSelect: run, icon: 'mail' }),
				command('send'),
				command('later', { group: 'bottom', disabled: ({ n }) => n > 1 }),
				command('never', { when: () => false }),
			],
			{ n: 2 },
		);
		expect(top.map((entry) => entry.id)).toEqual(['host:send']);
		expect(top[0]?.icon).toBe('mail');
		expect(bottom.map((entry) => [entry.id, entry.disabled])).toEqual([['host:later', true]]);
		top[0]?.onSelect();
		expect(run).toHaveBeenCalledWith({ n: 2 });
		expect(isHostMenuEntry(top[0])).toBe(true);
		expect(isHostMenuEntry({ id: 'copy' })).toBe(false);
	});

	it('never runs a disabled command', () => {
		const run = vi.fn();
		const { top } = resolveHostMenuCommands([command('a', { onSelect: run, disabled: true })], {
			n: 0,
		});
		top[0]?.onSelect();
		expect(run).not.toHaveBeenCalled();
	});

	it('separates the host groups from the built-in entries', () => {
		const menu = withHostMenuCommands(
			[{ id: 'cut' }, { id: 'copy' }] as { id: string; separatorBefore?: boolean }[],
			[command('top'), command('end', { group: 'bottom' })],
			{ n: 0 },
		);
		expect(menu.map((entry) => [entry.id, entry.separatorBefore ?? false])).toEqual([
			['host:top', false],
			['cut', true],
			['copy', false],
			['host:end', true],
		]);
		const alone = withHostMenuCommands<{ id: string; separatorBefore?: boolean }, { n: number }>(
			[],
			[command('end', { group: 'bottom' })],
			{ n: 0 },
		);
		expect(alone[0]?.separatorBefore).toBeUndefined();
	});
});

describe('regions', () => {
	it('compares renderers by function, tag or element spec', () => {
		const mount = () => undefined;
		expect(sameRenderer<unknown, unknown>(mount, mount)).toBe(true);
		expect(sameRenderer<unknown, unknown>('x-a', 'x-a')).toBe(true);
		expect(sameRenderer<unknown, unknown>('x-a', 'x-b')).toBe(false);
		expect(
			sameRenderer<unknown, unknown>(
				{ element: 'x-a', properties: { t: 1 } },
				{ element: 'x-a', properties: { t: 1 } },
			),
		).toBe(true);
		expect(
			sameRenderer<unknown, unknown>(
				{ element: 'x-a', properties: { t: 1 } },
				{ element: 'x-a', properties: { t: 2 } },
			),
		).toBe(false);
		expect(sameRenderer<unknown, unknown>(mount, undefined)).toBe(false);
	});

	it('compares contexts shallowly, arrays by value, skipping callbacks', () => {
		const base = { region: 'r', ids: ['a'], close: () => undefined };
		expect(sameRegionContext(base, { ...base, ids: ['a'], close: () => undefined })).toBe(true);
		expect(sameRegionContext(base, { ...base, ids: ['b'] })).toBe(false);
		expect(sameRegionContext(base, { ...base, region: 's' })).toBe(false);
		expect(
			sameRegionContext({ ...base, run: 1 }, { ...base, run: 2 }, new Set(['run', 'close'])),
		).toBe(true);
	});
});
