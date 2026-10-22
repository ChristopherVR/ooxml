import { afterEach, describe, expect, it, vi } from 'vitest';
import type { HostBackstagePage, HostRegionRenderer } from 'ooxml-core/embedding';
import { createHostBackstagePages } from './host-backstage';
import { withHostMenuItems } from './host-menu';
import { HOST_REGION_ATTRIBUTE, HostRegionHost } from './region-host';

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

interface Context {
	readonly region: string;
	readonly locale: string;
	readonly close: () => void;
}
const context = (locale = 'en', close = () => undefined): Context => ({
	region: 'backstage.host.a',
	locale,
	close,
});

describe('HostRegionHost', () => {
	it('mounts a function, updates it on a changed context and destroys it', () => {
		const container = document.createElement('div');
		const update = vi.fn();
		const destroy = vi.fn();
		const mount = vi.fn((target: HTMLElement) => {
			target.textContent = 'host';
			return { update, destroy };
		});
		const host = new HostRegionHost<Context>(container);
		host.render(mount, context());
		expect(container.textContent).toBe('host');
		expect(container.getAttribute(HOST_REGION_ATTRIBUTE)).toBe('backstage.host.a');
		host.render(
			mount,
			context('en', () => 1 as unknown as undefined),
		);
		expect(update).not.toHaveBeenCalled();
		host.render(mount, context('fr'));
		expect(update).toHaveBeenCalledWith(expect.objectContaining({ locale: 'fr' }));
		expect(mount).toHaveBeenCalledTimes(1);
		host.destroy();
		expect(destroy).toHaveBeenCalledTimes(1);
		expect(host.mounted).toBe(false);
	});

	it('creates a custom element, keeps its context current and calls the latest close', () => {
		const container = document.createElement('div');
		const host = new HostRegionHost<Context>(container);
		const first = vi.fn();
		const latest = vi.fn();
		host.render({ element: 'x-host-page', properties: { tenant: 't' } }, context('en', first));
		const element = container.firstElementChild as HTMLElement & {
			tenant: string;
			context: Context;
		};
		expect(element.localName).toBe('x-host-page');
		expect(element.tenant).toBe('t');
		const kept = element.context;
		host.render({ element: 'x-host-page', properties: { tenant: 't' } }, context('de', latest));
		expect(container.firstElementChild).toBe(element);
		expect(element.context.locale).toBe('de');
		kept.close();
		expect(latest).toHaveBeenCalledTimes(1);
		expect(first).not.toHaveBeenCalled();
		host.render('x-other', context('de', latest));
		expect(element.isConnected).toBe(false);
		expect(container.firstElementChild?.localName).toBe('x-other');
	});

	it('survives a renderer that throws', () => {
		const error = vi.spyOn(console, 'error').mockImplementation(() => undefined);
		const host = new HostRegionHost<Context>(document.createElement('div'));
		expect(() =>
			host.render(() => {
				throw new Error('boom');
			}, context()),
		).not.toThrow();
		expect(error).toHaveBeenCalled();
	});
});

describe('createHostBackstagePages', () => {
	type Page = HostBackstagePage<HostRegionRenderer<HTMLElement, Context>>;
	const page = (id: string, render: Page['render'], extra: Partial<Page> = {}): Page => ({
		id,
		label: id.toUpperCase(),
		render,
		...extra,
	});

	it('places nav items, mounts a page on show and tears removed pages down', () => {
		const backstage = document.createElement('div');
		document.body.append(backstage);
		const destroy = vi.fn();
		const pages = createHostBackstagePages<Context>(backstage, {
			context: (region) => ({ region, locale: 'en', close: () => undefined }),
			pageClass: 'page',
		});
		pages.setPages([
			page('about', 'x-about', { group: 'footer' }),
			page('policies', (target) => {
				target.textContent = 'Policies';
				return destroy;
			}),
		]);
		const items = pages.items([
			{ id: 'info', label: 'Info' },
			{ id: 'options', label: 'Options', group: 'footer' },
		]);
		expect(items.map((item) => item.id)).toEqual([
			'info',
			'host.policies',
			'options',
			'host.about',
		]);
		expect(items[3]).toEqual({ id: 'host.about', label: 'ABOUT', group: 'footer' });
		expect(pages.show('info')).toBe(false);
		expect(backstage.children).toHaveLength(0);
		expect(pages.show('host.policies')).toBe(true);
		const body = backstage.querySelector<HTMLElement>('[data-backstage-page="host.policies"]');
		expect(body?.textContent).toBe('Policies');
		expect(body?.className).toBe('page');
		expect(body?.getAttribute(HOST_REGION_ATTRIBUTE)).toBe('backstage.host.policies');
		pages.setPages([page('about', 'x-about')]);
		expect(destroy).toHaveBeenCalledTimes(1);
		expect(backstage.querySelector('[data-backstage-page="host.policies"]')).toBeNull();
		expect(pages.show('host.policies')).toBe(false);
		pages.show('host.about');
		pages.destroy();
		expect(backstage.children).toHaveLength(0);
	});
});

describe('withHostMenuItems', () => {
	it('draws host rows around the built-in ones and runs only host rows', () => {
		const run = vi.fn();
		const menu = withHostMenuItems(
			[{ id: 'copy', label: 'Copy' }],
			[
				{ id: 'send', label: 'Send', icon: 'mail', onSelect: run },
				{ id: 'off', label: 'Off', group: 'bottom', disabled: true, onSelect: run },
			],
			{ cell: 'A1' },
		);
		expect(menu.items).toEqual([
			{ id: 'host:send', label: 'Send', icon: 'mail' },
			{ id: 'copy', label: 'Copy', separatorBefore: true },
			{ id: 'host:off', label: 'Off', separatorBefore: true, disabled: true },
		]);
		expect(menu.run('copy')).toBe(false);
		expect(menu.run('host:off')).toBe(true);
		expect(run).not.toHaveBeenCalled();
		expect(menu.run('host:send')).toBe(true);
		expect(run).toHaveBeenCalledWith({ cell: 'A1' });
	});
});
