import { afterEach, expect, it, vi } from 'vitest';
import { mountTeams } from './index';

afterEach(() => document.body.replaceChildren());

it('mountTeams forwards className, props and (detail, event) open-file calls', () => {
	const container = document.createElement('div');
	document.body.append(container);
	const onOpenFile = vi.fn();
	const mounted = mountTeams(container, { workspaceId: 'acme', className: 'host', onOpenFile });
	expect(mounted.element.parentElement).toBe(container);
	expect(mounted.element.workspaceId).toBe('acme');
	expect(mounted.element.className).toBe('host');
	mounted.update({ workspaceId: 'acme', className: 'other', onOpenFile });
	expect(mounted.element.className).toBe('other');
	const detail = { name: 'a.xlsx' };
	const event = new CustomEvent('teams-open-file', { detail });
	mounted.element.dispatchEvent(event);
	expect(onOpenFile).toHaveBeenCalledWith(detail, event);
	mounted.destroy();
	expect(container.children.length).toBe(0);
});
