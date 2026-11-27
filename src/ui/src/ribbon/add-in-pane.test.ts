// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { openAddInPane } from './add-in-pane';
import { ADD_IN_PROTOCOL, ADD_IN_TIMEOUT } from './add-in-protocol';
import { ADD_IN_RUNTIME_CHANGED } from './add-in-runtime-settings';

const channels: MessageChannel[] = [];
afterEach(() => {
	document.body.replaceChildren();
	for (const channel of channels.splice(0)) {
		channel.port1.close();
		channel.port2.close();
	}
	vi.useRealTimers();
});

function mount(
	host = document.body.appendChild(document.createElement('div')),
	active = () => true,
) {
	const restoreFocus = vi.fn();
	const readText = vi.fn(() => 'selected');
	openAddInPane(host, {
		id: 'notes',
		url: new URL('https://publisher.example.test/pane'),
		name: 'Notes',
		product: 'word',
		permissions: ['Document.Read.User'],
		selection: { readText, writeText: () => false },
		active,
		scopes: [host],
		t: (text) => text,
		restoreFocus,
	});
	const pane = host.querySelector('office-ui-add-in-pane');
	const frame = pane?.shadowRoot?.querySelector('iframe');
	if (!pane || !frame) throw new Error('Missing pane');
	return { host, pane, frame, restoreFocus, readText };
}

function connect(frame: HTMLIFrameElement) {
	const channel = new MessageChannel();
	channels.push(channel);
	Object.defineProperty(frame, 'contentWindow', { value: window });
	const post = vi.spyOn(channel.port1, 'postMessage');
	const close = vi.spyOn(channel.port1, 'close');
	window.dispatchEvent(
		new MessageEvent('message', {
			source: window,
			origin: 'null',
			ports: [channel.port1],
			data: { kind: 'connect', protocol: ADD_IN_PROTOCOL },
		}),
	);
	expect(post).toHaveBeenCalledWith({ kind: 'ready', host: 'Word', hasSelection: true });
	return {
		post,
		close,
		request: (id: number) => channel.port2.postMessage({ id, method: 'getSelectedDataAsync' }),
	};
}

it('removes an unconnected frame on timeout and reports it', () => {
	vi.useFakeTimers();
	const { pane, frame } = mount();
	vi.advanceTimersByTime(ADD_IN_TIMEOUT);
	expect(frame.isConnected).toBe(false);
	expect(pane.shadowRoot?.querySelector('[role="status"]')?.textContent).toContain(
		'did not connect',
	);
});

it('closes the port and removes the frame on its second load', () => {
	const { pane, frame } = mount();
	const { close } = connect(frame);
	frame.dispatchEvent(new Event('load'));
	expect(frame.isConnected).toBe(true);
	frame.dispatchEvent(new Event('load'));
	expect(close).toHaveBeenCalledOnce();
	expect(frame.isConnected).toBe(false);
	expect(pane.shadowRoot?.querySelector('[role="status"]')?.textContent).toContain('disconnected');
});

it('ignores replayed ids and closes after the 256th accepted request', async () => {
	const { frame, readText } = mount();
	const { post, close, request } = connect(frame);
	for (let id = 1; id <= 256; id++) request(id);
	request(256);
	request(255);
	request(257);
	await vi.waitFor(() => expect(close).toHaveBeenCalledOnce());
	expect(readText).toHaveBeenCalledTimes(256);
	expect(post).toHaveBeenCalledTimes(258);
	expect(post).toHaveBeenLastCalledWith({
		id: 257,
		error: expect.stringContaining('request limit'),
	});
});

it('refuses requests from a revoked install and removes its pane when notified', async () => {
	let active = true;
	const { host, pane, frame, readText, restoreFocus } = mount(undefined, () => active);
	const { post, close, request } = connect(frame);
	active = false;
	request(1);
	await vi.waitFor(() =>
		expect(post).toHaveBeenLastCalledWith({
			id: 1,
			error: expect.stringContaining('no longer has permission'),
		}),
	);
	expect(readText).not.toHaveBeenCalled();
	host.dispatchEvent(new Event(ADD_IN_RUNTIME_CHANGED));
	expect(pane.isConnected).toBe(false);
	expect(close).toHaveBeenCalledOnce();
	expect(restoreFocus).toHaveBeenCalledOnce();
});

it('closes the preceding channel when another pane replaces it or the pane is removed', () => {
	const first = mount();
	const firstConnection = connect(first.frame);
	const replacement = mount(first.host);
	expect(first.pane.isConnected).toBe(false);
	expect(firstConnection.close).toHaveBeenCalledOnce();
	const replacementConnection = connect(replacement.frame);
	replacement.pane.remove();
	expect(replacementConnection.close).toHaveBeenCalledOnce();
});

it('removes a disconnected pane when its host reconnects', () => {
	const { host, pane, frame } = mount();
	const { close } = connect(frame);
	host.remove();
	document.body.append(host);
	expect(close).toHaveBeenCalledOnce();
	expect(host.querySelector('office-ui-add-in-pane')).toBeNull();
});
