import { describe, expect, it, vi } from 'vitest';
import { bindTeams, defineTeamsApp, pickTeamsProps, type TeamsApp } from './index';

function element(): TeamsApp {
	defineTeamsApp();
	return document.createElement('teams-app') as TeamsApp;
}

describe('pickTeamsProps', () => {
	it('keeps provided values, an explicit null config, and drops undefined and handlers', () => {
		expect(
			pickTeamsProps({ workspaceId: 'acme', userName: undefined, config: null, className: 'a b' }),
		).toEqual({ workspaceId: 'acme', config: null, className: 'a b' });
		expect(pickTeamsProps({ className: '' })).toEqual({});
	});
});

describe('bindTeams', () => {
	it('applies props, swaps the host class list and re-applies only what changed', () => {
		const el = element();
		const binding = bindTeams(el, () => ({}));
		binding.update({ workspaceId: 'acme', userName: 'Ada', className: 'one two' });
		expect(el.workspaceId).toBe('acme');
		expect(el.userName).toBe('Ada');
		expect(el.className).toBe('one two');
		el.userName = 'edited in the element';
		binding.update({ workspaceId: 'acme', userName: 'Ada', className: 'two three' });
		expect(el.userName).toBe('edited in the element');
		expect([...el.classList]).toEqual(['two', 'three']);
		binding.update({ workspaceId: 'acme', userName: 'Ada' });
		expect(el.classList.length).toBe(0);
		binding.destroy();
	});

	it('calls the latest handlers with (detail, event) for open-file and stops on destroy', () => {
		const el = element();
		const first = vi.fn();
		const latest = vi.fn();
		let handlers = { onOpenFile: first };
		const binding = bindTeams(el, () => handlers);
		handlers = { onOpenFile: latest };
		const detail = { name: 'a.xlsx' };
		const event = new CustomEvent('teams-open-file', { detail, cancelable: true });
		el.dispatchEvent(event);
		expect(first).not.toHaveBeenCalled();
		expect(latest).toHaveBeenCalledWith(detail, event);
		binding.destroy();
		el.dispatchEvent(new CustomEvent('teams-open-file', { detail }));
		expect(latest).toHaveBeenCalledOnce();
	});
});
