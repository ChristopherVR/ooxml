// @vitest-environment jsdom
import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot, createSignal } from 'solid-js';
import { createApp, h, nextTick, ref } from 'vue';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
	bindElement,
	contractOptions,
	handleKeys,
	readOnlyChange,
	type BindingContract,
} from './index';
import { createReactBinding } from './react';
import { createSolidBinding } from './solid';
import { defineVueBinding } from './vue';

// A minimal product for the framework factory tests: one element, two props, two events and a
// small handle, wired through the kit the way a real product's mount function is.

interface DemoProps {
	label?: string | undefined;
	readOnly?: boolean | undefined;
}
interface DemoOptions extends DemoProps {
	onPing?: ((value: number) => void) | undefined;
	onReadOnlyChange?: ((readOnly: boolean, detail: { hostReadOnly: boolean }) => void) | undefined;
}
interface DemoHandle {
	readonly element: HTMLElement;
	ping(value: number): void;
	readonly dirty: boolean;
}
interface DemoBinding extends DemoHandle {
	update(options: DemoOptions): void;
	destroy(): void;
}

const DEMO_CONTRACT: BindingContract<HTMLElement, DemoProps, DemoHandle> = {
	props: {
		label: { forward: 'changed', apply: (element, value) => (element.title = value ?? '') },
		readOnly: {
			forward: 'changed',
			fallback: false,
			apply: (element, value) => element.toggleAttribute('read-only', value === true),
		},
	},
	events: { ping: { option: 'onPing' }, 'readonly-change': readOnlyChange },
	handle: { element: 'get', ping: 'call', dirty: { fallback: false } },
};

const updates = vi.fn();
const destroyed = vi.fn();

function mountDemo(host: HTMLElement, initial: DemoOptions): DemoBinding {
	const element = host.ownerDocument.createElement('demo-element');
	let options = initial;
	const bound = bindElement(element, DEMO_CONTRACT, () => options);
	const binding: DemoBinding = {
		element,
		update(next) {
			options = next;
			updates(next);
			bound.update(next);
		},
		ping: (value) => element.dispatchEvent(new CustomEvent('ping', { detail: value })),
		get dirty() {
			return element.hasAttribute('dirty');
		},
		destroy() {
			destroyed();
			bound.unlisten();
			element.remove();
		},
	};
	binding.update(initial);
	host.append(element);
	return binding;
}

function emitReadOnly(element: HTMLElement, hostReadOnly: boolean): void {
	element.dispatchEvent(
		new CustomEvent('readonly-change', {
			detail: { readOnly: true, locked: !hostReadOnly, hostReadOnly },
		}),
	);
}

const reactAct = globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean };
beforeEach(() => {
	reactAct.IS_REACT_ACT_ENVIRONMENT = true;
	updates.mockClear();
	destroyed.mockClear();
});
afterEach(() => {
	delete reactAct.IS_REACT_ACT_ENVIRONMENT;
	document.body.replaceChildren();
});
const newHost = () => document.body.appendChild(document.createElement('div'));
const missing = (target: object) =>
	handleKeys(DEMO_CONTRACT.handle).filter((key) => !(key in target));

describe('createReactBinding', () => {
	const Demo = createReactBinding<DemoOptions & { className?: string }, DemoHandle>({
		displayName: 'Demo',
		handle: DEMO_CONTRACT.handle,
		mount: mountDemo,
	});
	it('mounts once, forwards props, exposes the handle and unmounts', () => {
		const root = createRoot(newHost());
		const handle = createRef<DemoHandle>();
		const onPing = vi.fn();
		act(() =>
			root.render(createElement(Demo, { ref: handle, label: 'a', className: 'host', onPing })),
		);
		const element = document.querySelector('demo-element') as HTMLElement;
		expect(element.title).toBe('a');
		expect(element.parentElement?.className).toBe('host');
		expect(missing(handle.current!)).toEqual([]);
		expect(handle.current!.element).toBe(element);
		handle.current!.ping(3);
		expect(onPing).toHaveBeenCalledWith(3);
		act(() => root.render(createElement(Demo, { ref: handle, label: 'b', onPing })));
		expect(element.title).toBe('b');
		expect(Demo.displayName).toBe('Demo');
		act(() => root.unmount());
		expect(destroyed).toHaveBeenCalledOnce();
		expect(handle.current).toBeNull();
	});
});

describe('defineVueBinding', () => {
	const Demo = defineVueBinding({
		name: 'Demo',
		props: { label: String, readOnly: Boolean },
		contract: DEMO_CONTRACT,
		handle: DEMO_CONTRACT.handle,
		mount: mountDemo,
	});
	it('declares the emits, re-emits events, mirrors v-model:readOnly and exposes the handle', async () => {
		const label = ref('a');
		const readOnly = ref(false);
		const editor = ref<DemoHandle>();
		const onPing = vi.fn();
		const onUpdate = vi.fn((value: boolean) => (readOnly.value = value));
		const onReadonly = vi.fn();
		const app = createApp({
			render: () =>
				h(Demo, {
					ref: editor,
					label: label.value,
					readOnly: readOnly.value,
					onPing,
					'onReadonly-change': onReadonly,
					'onUpdate:readOnly': onUpdate,
				}),
		});
		expect(Demo.emits).toEqual(['ping', 'readonly-change', 'update:readOnly']);
		app.mount(newHost());
		const element = document.querySelector('demo-element') as HTMLElement;
		expect(missing(editor.value!)).toEqual([]);
		editor.value!.ping(2);
		expect(onPing).toHaveBeenCalledWith(2);
		emitReadOnly(element, false);
		expect(onReadonly).toHaveBeenCalledOnce();
		expect(onUpdate).not.toHaveBeenCalled();
		emitReadOnly(element, true);
		expect(onUpdate).toHaveBeenCalledWith(true);
		await nextTick();
		expect(element.hasAttribute('read-only')).toBe(true);
		label.value = 'b';
		await nextTick();
		expect(element.title).toBe('b');
		app.unmount();
		expect(destroyed).toHaveBeenCalledOnce();
		expect(editor.value).toBeNull();
	});
	it('reads the element leniently before mount', () => {
		const Lenient = defineVueBinding({
			name: 'Lenient',
			props: {},
			contract: DEMO_CONTRACT,
			handle: DEMO_CONTRACT.handle,
			mount: mountDemo,
		});
		expect(Lenient.name).toBe('Lenient');
	});
});

describe('createSolidBinding', () => {
	const Demo = createSolidBinding<
		DemoOptions & { class?: string; editorRef?: (handle: DemoHandle) => void },
		DemoHandle,
		DemoOptions
	>({
		handle: DEMO_CONTRACT.handle,
		mount: mountDemo,
		options: (props) => contractOptions(DEMO_CONTRACT, props),
	});
	it('mounts once, reacts to props, hands over the handle and cleans up', () => {
		let handle: DemoHandle | undefined;
		const [label, setLabel] = createSignal('a');
		let dispose!: () => void;
		let host!: HTMLDivElement;
		createSolidRoot((stop) => {
			dispose = stop;
			host = Demo({
				get label() {
					return label();
				},
				class: 'host',
				editorRef: (value) => (handle = value),
			});
		});
		document.body.append(host);
		const element = host.querySelector('demo-element') as HTMLElement;
		expect(host.className).toBe('host');
		expect(element.title).toBe('a');
		expect(missing(handle!)).toEqual([]);
		setLabel('b');
		expect(element.title).toBe('b');
		expect(updates.mock.calls.length).toBeGreaterThanOrEqual(2);
		dispose();
		expect(destroyed).toHaveBeenCalledOnce();
	});
});
