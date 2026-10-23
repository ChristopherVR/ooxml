// Every framework adapter against `VIEWER_CONTRACT`: the same properties in, every event out
// (through the `events` callback map and the native output), and the same handle members.
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import '@angular/compiler';
import { reflectComponentType } from '@angular/core';
import { handleKeys, propKeys } from 'ooxml-ui/bindings';
import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { createApp, h, ref } from 'vue';
import { current, reset } from './mock-binding';
import { VisioViewerComponent } from '../src/angular';
import { eventKeys, propertyKeys, VIEWER_CONTRACT, type ViewerHandle } from '../src/common';
import { VisioViewer as ReactViewer } from '../src/react';
import { VisioViewer as SolidViewer } from '../src/solid';
import { VisioViewer as VueViewer } from '../src/vue';

beforeEach(reset);
const sorted = (values: Iterable<string>) => [...values].sort();
const HANDLE = handleKeys(VIEWER_CONTRACT.handle);
const missing = (target: object) => HANDLE.filter((key) => !(key in target));
const newHost = () => document.body.appendChild(document.createElement('div'));
/** Calls every event through the mounted binding's callback map; returns the events that ran. */
function fireAll(callbacks: Record<string, ReturnType<typeof vi.fn>>): string[] {
	for (const name of eventKeys) current().options.events?.[name]?.(undefined as never);
	return Object.keys(callbacks).filter((name) => callbacks[name]!.mock.calls.length > 0);
}
const spies = () => Object.fromEntries(eventKeys.map((name) => [name, vi.fn()]));

describe('VIEWER_CONTRACT', () => {
	it('lists every shared property and the members of ViewerHandle', () => {
		expect(sorted(propKeys(VIEWER_CONTRACT))).toEqual(sorted(propertyKeys));
		expect(HANDLE).toContain('element');
		expect(HANDLE).toContain('createPrintSnapshot');
	});
});

describe('every framework adapter matches the contract', () => {
	it('React: forwards every event and exposes the handle', async () => {
		const root = createRoot(newHost());
		const handle = createRef<ViewerHandle>();
		const events = spies();
		await act(async () => root.render(createElement(ReactViewer, { ref: handle, events })));
		expect(missing(handle.current!)).toEqual([]);
		expect(sorted(fireAll(events))).toEqual(sorted(eventKeys));
		await act(async () => root.unmount());
	});
	it('Vue: declares every property, emits every event and exposes the handle', () => {
		expect(sorted(Object.keys(VueViewer.props))).toEqual(sorted([...propertyKeys, 'events']));
		expect(sorted(VueViewer.emits as string[])).toEqual(sorted(eventKeys));
		const viewer = ref<ViewerHandle>();
		const events = spies();
		const app = createApp({ render: () => h(VueViewer, { ref: viewer, events }) });
		app.mount(newHost());
		expect(missing(viewer.value!)).toEqual([]);
		expect(sorted(fireAll(events))).toEqual(sorted(eventKeys));
		app.unmount();
	});
	it('Solid: forwards every event and hands viewerRef the handle', () => {
		let handle: ViewerHandle | undefined;
		const events = spies();
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			newHost().append(
				SolidViewer({ events, viewerRef: (value) => (handle = value) }) as unknown as Node,
			);
		});
		expect(missing(handle!)).toEqual([]);
		expect(sorted(fireAll(events))).toEqual(sorted(eventKeys));
		dispose();
	});
	it('Angular: one input per property, one output per event, the handle on the class', () => {
		const mirror = reflectComponentType(VisioViewerComponent)!;
		expect(sorted(mirror.inputs.map((input) => input.propName))).toEqual(
			sorted([...propertyKeys, 'events']),
		);
		expect(mirror.outputs).toHaveLength(eventKeys.length);
		expect(missing(VisioViewerComponent.prototype)).toEqual([]);
	});
	it('Svelte: one prop per property and every handle method exported', () => {
		const source = readFileSync(resolve(__dirname, '../src/VisioViewer.svelte'), 'utf8');
		const declared = /let \{([^}]*)\}: Props = \$props\(\)/u.exec(source)?.[1] ?? '';
		const names = declared.split(',').map((entry) => entry.split(':')[0]!.trim());
		expect(sorted(names)).toEqual(
			sorted([...propertyKeys, 'events', 'class', "'aria-label'", 'style']),
		);
		// Svelte exports methods only; `element` and `controller` come from `getHandle()`.
		for (const key of HANDLE.filter((name) => name !== 'element' && name !== 'controller'))
			expect(source, key).toMatch(new RegExp(String.raw`export function ${key}\(`));
		expect(source).toMatch(/export function getHandle\(/u);
	});
});
