import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	angularOutputs,
	bindElement,
	contractOptions,
	createPropTracker,
	createSrcLoader,
	deferredHandle,
	eventArgs,
	eventOptions,
	handleKeys,
	installHandle,
	listenEvents,
	nameFromUrl,
	pickProps,
	readOnlyChange,
	sameList,
	shallowEqual,
	svelteHandlerName,
	syncClassName,
	type BindingContract,
	type HandleSpec,
} from './index';
import { vueEmits } from './vue';

afterEach(() => {
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});

class FakeElement extends EventTarget {
	locale = 'en';
	readOnly = false;
	fileName = 'Book1';
	colors: Record<string, string> = {};
	model: unknown;
	writes: string[] = [];
}
interface FakeProps {
	locale?: string | undefined;
	readOnly?: boolean | undefined;
	fileName?: string | undefined;
	colors?: Record<string, string> | undefined;
	model?: unknown;
}
const contract: BindingContract<
	FakeElement,
	FakeProps,
	unknown,
	'change' | 'readonly-change' | 'file-command'
> = {
	props: {
		locale: { forward: 'always', fallback: 'en' },
		readOnly: { forward: 'changed', fallback: false },
		fileName: { forward: 'defined' },
		colors: {
			forward: 'changed',
			same: shallowEqual,
			apply: (element, value) => {
				element.writes.push('colors');
				element.colors = { ...value };
			},
		},
		model: { forward: 'manual' },
	},
	events: {
		change: { option: 'onChange' },
		'readonly-change': readOnlyChange,
		'file-command': { option: 'onFileCommand', args: eventArgs.event },
	},
	handle: {},
};

describe('equality and names', () => {
	it('compares lists and objects by value, treating missing as empty', () => {
		expect(sameList(['a'], ['a'])).toBe(true);
		expect(sameList(undefined, [])).toBe(true);
		expect(sameList(['a'], ['a', 'b'])).toBe(false);
		expect(shallowEqual({ a: 1 }, { a: 1 })).toBe(true);
		expect(shallowEqual(undefined, {})).toBe(true);
		expect(shallowEqual({ a: 1 }, { a: 1, b: 2 })).toBe(false);
	});
	it('names a file after the last URL segment', () => {
		expect(nameFromUrl('/a/Q3%20Sales.xlsx#top')).toBe('Q3 Sales.xlsx');
		expect(nameFromUrl('/a/bad%zz.xlsx?x=1')).toBe('bad%zz.xlsx');
		expect(nameFromUrl('/a/')).toBeUndefined();
	});
	it('derives framework handler names from the options', () => {
		expect(svelteHandlerName('onRibbonAddIn')).toBe('onribbonaddin');
	});
});

describe('createPropTracker', () => {
	it('reports the first value and then only changes, and remembers seeds', () => {
		const changed = createPropTracker<'a' | 'b'>();
		expect(changed('a', 1)).toBe(true);
		expect(changed('a', 1)).toBe(false);
		expect(changed('a', 2)).toBe(true);
		expect(changed('a', [1], sameList)).toBe(true);
		expect(changed('a', [1], sameList)).toBe(false);
		changed.seed('b', null);
		expect(changed('b', null)).toBe(false);
		expect(changed.sent('a')).toEqual([1]);
	});
});

describe('bindElement', () => {
	it('forwards each prop by its policy and leaves manual props to the product', () => {
		const element = new FakeElement();
		const bound = bindElement(element, contract, () => ({}));
		bound.update({ fileName: 'A', colors: { x: '1' }, model: 'm' });
		expect(element).toMatchObject({ locale: 'en', fileName: 'A', colors: { x: '1' } });
		expect(element.model).toBeUndefined();
		// What the user changes inside the element survives unrelated re-renders.
		element.readOnly = true;
		element.fileName = 'Opened';
		element.locale = 'de';
		bound.update({ readOnly: false, colors: { x: '1' } });
		expect(element.readOnly).toBe(true);
		expect(element.fileName).toBe('Opened');
		// `always` is written every time; an inline object equal by value is not a change.
		expect(element.locale).toBe('en');
		expect(element.writes).toEqual(['colors']);
		bound.update({ readOnly: true, fileName: 'A' });
		expect(element.fileName).toBe('Opened');
		bound.update({ readOnly: false, fileName: 'B' });
		expect(element).toMatchObject({ readOnly: false, fileName: 'B' });
		expect(bound.sent('fileName')).toBe('B');
		bound.unlisten();
	});
	it('treats an initial value as already written', () => {
		const element = new FakeElement();
		const set = vi.fn();
		const bound = bindElement(
			element,
			{
				props: {
					readOnly: { forward: 'changed' as const, initial: undefined, apply: set },
				} as BindingContract<FakeElement, { readOnly?: boolean | undefined }>['props'],
				events: {},
			},
			() => ({}),
		);
		bound.update({});
		expect(set).not.toHaveBeenCalled();
		bound.update({ readOnly: true });
		expect(set).toHaveBeenCalledWith(element, true);
	});
	it('calls the current handlers with each rule’s arguments, after onEvent', () => {
		const element = new FakeElement();
		const order: string[] = [];
		let handlers: Record<string, unknown> = { onChange: () => order.push('old') };
		const bound = bindElement(element, contract, () => handlers, {
			onEvent: (name) => order.push(`seen ${name}`),
		});
		const onChange = vi.fn(() => order.push('handler'));
		const onReadOnlyChange = vi.fn();
		const onFileCommand = vi.fn();
		handlers = { onChange, onReadOnlyChange, onFileCommand };
		element.dispatchEvent(new CustomEvent('change', { detail: 7 }));
		const detail = { readOnly: true, locked: true, hostReadOnly: false };
		element.dispatchEvent(new CustomEvent('readonly-change', { detail }));
		const command = new CustomEvent('file-command', { detail: { command: 'save' } });
		element.dispatchEvent(command);
		expect(order).toEqual(['seen change', 'handler', 'seen readonly-change', 'seen file-command']);
		expect(onChange).toHaveBeenCalledWith(7);
		expect(onReadOnlyChange).toHaveBeenCalledWith(true, detail);
		expect(onFileCommand).toHaveBeenCalledWith(command);
		bound.unlisten(['change']);
		element.dispatchEvent(new CustomEvent('change', { detail: 8 }));
		expect(onChange).toHaveBeenCalledOnce();
		element.dispatchEvent(new CustomEvent('readonly-change', { detail }));
		expect(onReadOnlyChange).toHaveBeenCalledTimes(2);
		bound.unlisten();
		element.dispatchEvent(new CustomEvent('readonly-change', { detail }));
		expect(onReadOnlyChange).toHaveBeenCalledTimes(2);
	});
	it('listens without a contract for products that only need events', () => {
		const element = new EventTarget();
		const onOpen = vi.fn();
		const stop = listenEvents(
			element,
			{ open: { option: 'onOpen', args: eventArgs.detailAndEvent } },
			() => ({ onOpen }),
		);
		const event = new CustomEvent('open', { detail: 'a' });
		element.dispatchEvent(event);
		expect(onOpen).toHaveBeenCalledWith('a', event);
		stop();
	});
	it('swaps only the classes the binding added', () => {
		const element = document.createElement('div');
		element.className = 'own';
		syncClassName(element, undefined, 'a b');
		syncClassName(element, 'a b', 'b c');
		expect([...element.classList]).toEqual(['own', 'b', 'c']);
		syncClassName(element, 'b c', undefined);
		expect(element.className).toBe('own');
	});
});

describe('contract mapping', () => {
	it('picks the props, maps handlers and reads framework-named handlers', () => {
		expect(pickProps(contract, { locale: 'fr', other: 1 })).toEqual({
			locale: 'fr',
			readOnly: undefined,
			fileName: undefined,
			colors: undefined,
			model: undefined,
		});
		expect(pickProps(contract, { locale: 'fr', fileName: undefined }, true)).toEqual({
			locale: 'fr',
		});
		const change = () => {};
		expect(eventOptions(contract, { change })).toEqual({
			onChange: change,
			onReadOnlyChange: undefined,
			onFileCommand: undefined,
		});
		const ready = () => {};
		expect(
			contractOptions(contract, { onchange: change, onready: ready }, svelteHandlerName, [
				'onReady',
			]),
		).toMatchObject({ onChange: change, onReady: ready });
	});
	it('emits Angular outputs, and a modelled event only when the input differs', () => {
		const component = {
			readOnly: false,
			change: { emit: vi.fn() },
			readOnlyChange: { emit: vi.fn() },
			fileCommand: { emit: vi.fn() },
			ready: { emit: vi.fn() },
		};
		const options = angularOutputs<Record<string, (...args: unknown[]) => void>>(
			contract,
			component,
			['onReady'],
		);
		options.onChange?.(1);
		options.onReady?.('el');
		options.onReadOnlyChange?.(true, { hostReadOnly: false });
		expect(component.readOnlyChange.emit).not.toHaveBeenCalled();
		options.onReadOnlyChange?.(true, { hostReadOnly: true });
		expect(component.readOnlyChange.emit).toHaveBeenCalledWith(true);
		expect(component.change.emit).toHaveBeenCalledWith(1);
		expect(component.ready.emit).toHaveBeenCalledWith('el');
	});
	it('generates Vue emits with the model and the extra callbacks', () => {
		expect(vueEmits(contract, { ready: 'onReady' })).toEqual([
			'change',
			'readonly-change',
			'file-command',
			'update:readOnly',
			'ready',
		]);
	});
});

interface Handle {
	readonly element: string;
	load(input: number): Promise<number>;
	select(ref: string): string;
	markClean(): void;
	readonly dirty: boolean;
}
const spec: HandleSpec<Handle> = {
	element: 'get',
	load: 'async',
	select: 'call',
	markClean: 'optional-call',
	dirty: { fallback: false },
};
const binding = (): Handle => ({
	element: 'el',
	load: async (input) => input * 2,
	select: (ref) => `selected ${ref}`,
	markClean: vi.fn(),
	dirty: true,
});

describe('deferredHandle', () => {
	it('throws, rejects or falls back before mount and forwards after', async () => {
		let current: Handle | undefined;
		const handle = deferredHandle(spec, () => current, { message: 'The viewer is not mounted.' });
		expect(handleKeys(spec)).toEqual(['element', 'load', 'select', 'markClean', 'dirty']);
		expect(Object.keys(handle)).toEqual(handleKeys(spec));
		expect(() => handle.element).toThrow('The viewer is not mounted.');
		expect(() => handle.select('A1')).toThrow('not mounted');
		await expect(handle.load(1)).rejects.toThrow('not mounted');
		expect(handle.markClean()).toBeUndefined();
		expect(handle.dirty).toBe(false);
		current = binding();
		expect(handle.element).toBe('el');
		await expect(handle.load(2)).resolves.toBe(4);
		expect(handle.select('B2')).toBe('selected B2');
		handle.markClean();
		expect(current.markClean).toHaveBeenCalledOnce();
		expect(handle.dirty).toBe(true);
	});
	it('reads getters leniently for Vue templates', () => {
		const handle = deferredHandle(spec, () => undefined, { lenientGetters: true });
		expect(handle.element).toBeUndefined();
		expect(() => handle.select('A1')).toThrow('Editor is not mounted');
	});
	it('installs the members on a class prototype', async () => {
		class Component {
			binding?: Handle;
		}
		installHandle(Component.prototype, spec, (self) => self.binding, { lenientGetters: true });
		const instance = new Component() as Component & Handle;
		expect('select' in Component.prototype).toBe(true);
		expect(instance.element).toBeUndefined();
		instance.binding = binding();
		expect(instance.element).toBe('el');
		await expect(instance.load(3)).resolves.toBe(6);
	});
});

describe('createSrcLoader', () => {
	it('loads under the URL name, reports failures and drops superseded fetches', async () => {
		const resolvers: ((response: Response) => void)[] = [];
		vi.stubGlobal(
			'fetch',
			vi.fn(() => new Promise<Response>((resolve) => resolvers.push(resolve))),
		);
		const load = vi.fn(async () => {});
		const report = vi.fn();
		const loader = createSrcLoader(load, report);
		loader.fetch('/files/first.xlsx', undefined);
		await Promise.resolve();
		loader.fetch('/files/Second%20one.xlsx', undefined);
		await Promise.resolve();
		resolvers[0]?.(new Response(new Uint8Array([1])));
		resolvers[1]?.(new Response(new Uint8Array([2])));
		await vi.waitFor(() => expect(load).toHaveBeenCalledOnce());
		expect(load).toHaveBeenCalledWith(new Uint8Array([2]), 'Second one.xlsx');
		loader.fetch('/missing.xlsx', 'Named.xlsx');
		await Promise.resolve();
		resolvers[2]?.(new Response('', { status: 404 }));
		await vi.waitFor(() => expect(report).toHaveBeenCalledOnce());
		expect(String(report.mock.calls[0]?.[0])).toContain('HTTP 404');
		loader.fetch('/late.xlsx', undefined);
		await Promise.resolve();
		loader.cancel();
		resolvers[3]?.(new Response(new Uint8Array([3])));
		await new Promise((done) => setTimeout(done, 0));
		expect(load).toHaveBeenCalledOnce();
		expect(report).toHaveBeenCalledOnce();
	});
});
