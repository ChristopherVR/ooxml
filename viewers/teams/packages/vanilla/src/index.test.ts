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

// Region keys come from the shared element, so every binding that mounts it gets them.
it('F6 moves the keyboard from the composer round the workspace and back', async () => {
	const container = document.createElement('div');
	document.body.append(container);
	const mounted = mountTeams(container, {
		workspaceId: 'vanilla-regions',
		userName: 'Ada',
		userId: 'ada',
		config: { mode: 'local', iceServers: [] },
	});
	const element = mounted.element;
	await element.updateComplete;
	element.client!.createChannel('Ops');
	await new Promise((done) => setTimeout(done, 10));
	await element.updateComplete;
	const root = element.shadowRoot!;
	const deepActive = (): Element | null => {
		let active = root.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		return active;
	};
	const message = root
		.querySelector('office-ui-chat-composer')!
		.shadowRoot!.querySelector<HTMLTextAreaElement>('textarea')!;
	message.focus();
	const press = () =>
		deepActive()?.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'F6', bubbles: true, composed: true }),
		);
	press();
	expect(root.activeElement).toBe(root.querySelector('office-ui-app-rail'));
	for (let step = 0; step < 4; step += 1) press();
	expect(deepActive()).toBe(message);
	mounted.destroy();
});
